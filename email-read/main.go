package main

import (
	"fmt"
	"io"
	"io/ioutil"
	"log"

	"github.com/emersion/go-imap"
	"github.com/emersion/go-imap/client"
	"github.com/emersion/go-message/mail"
)

func connect() (*client.Client, error) {
	c, err := client.DialTLS(config.Imap, nil)

	if err != nil {
		log.Fatal(err)
	}
	log.Println("Connected")

	// Login
	if err := c.Login(config.Username, config.Password); err != nil {
		log.Fatal(err)
	}
	log.Println("Logged in")

	return c, err
}

func main() {
	log.Println("Connecting to server...")
	// conectar
	c, err := connect()
	// Don't forget to logout
	defer c.Logout()

	done := make(chan error, 1)
	// seleciona INBOX
	mbox, err := c.Select("INBOX", true)
	if err != nil {
		log.Fatal(err)
	}
	log.Println("Flags for INBOX:", mbox.Flags)

	// filtra mensagens não lidas
	criteria := imap.NewSearchCriteria()
	criteria.WithoutFlags = []string{"\\Seen"}
	uids, err := c.Search(criteria)

	seqset := new(imap.SeqSet)
	seqset.AddNum(uids...)

	messages := make(chan *imap.Message)
	done = make(chan error, 1)

	section := &imap.BodySectionName{}
	go func() {
		items := []imap.FetchItem{imap.FetchEnvelope, imap.FetchFlags, imap.FetchInternalDate, section.FetchItem()}
		done <- c.Fetch(seqset, items, messages)
	}()

	file := CreateFile()
	file.AddRow("Planilha1", []string{"Date", "From", "To", "Subject", "Content"})

	// mensagens
	for msg := range messages {
		log.Println("* "+msg.Envelope.Subject, msg.Flags)

		r := msg.GetBody(section)
		if r == nil {
			log.Fatal("Server didn't returned message body")
		}

		// Create a new mail reader
		mr, err := mail.CreateReader(r)
		if err != nil {
			log.Fatal(err)
		}

		var row []string

		// Print some info about the message
		header := mr.Header
		if date, err := header.Date(); err == nil {
			row = append(row, date.Format("2006-01-02 15:04:05"))
			log.Println("Date:", date)
		}
		if from, err := header.AddressList("From"); err == nil {
			row = append(row, from[0].Address)
			log.Println("From:", from)
		}
		if to, err := header.AddressList("To"); err == nil {
			row = append(row, to[0].Address)
			log.Println("To:", to)
		}
		if subject, err := header.Subject(); err == nil {
			row = append(row, subject)
			log.Println("Subject:", subject)
		}

		// Process each message's part
		for {
			p, err := mr.NextPart()
			if err == io.EOF {
				break
			} else if err != nil {
				log.Fatal(err)
			}

			switch h := p.Header.(type) {
			case *mail.InlineHeader:
				// This is the message's text (can be plain-text or HTML)
				b, _ := ioutil.ReadAll(p.Body)
				fmt.Println("Got text: %v", string(b))
				row = append(row, string(b))
			case *mail.AttachmentHeader:
				// This is an attachment
				filename, _ := h.Filename()
				fmt.Println("Got attachment: %v", filename)
			}
		}

		file.AddRow("Planilha1", row)
	}

	if err := <-done; err != nil {
		log.Fatal(err)
	}

	file.Save("teste-excel.xlsx")

	log.Println("Done!")
}
