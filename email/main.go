package main

import (
	"bufio"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"strconv"
	"strings"
	"time"

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
	file.AddRow("Planilha1", []string{"Data", "Pauta", "Veiculo", "Link"})

	fmt.Println("Lendo (" + strconv.Itoa(len(uids)) + ") emails")

	i := 1
	// mensagens
	for msg := range messages {

		fmt.Println("Lendo (" + strconv.Itoa(i) + " de " + strconv.Itoa(len(uids)) + ") emails")
		i++
		r := msg.GetBody(section)
		if r == nil {
			log.Fatal("Server didn't returned message body")
		}

		// Create a new mail reader
		mr, err := mail.CreateReader(r)
		if err != nil {
			continue
		}

		// Print some info about the message
		header := mr.Header
		date, _ := header.Date()

		//row = append(row, date.Format("2006-01-02 15:04:05"))

		if from, err := header.AddressList("From"); err == nil {
			if from[0].Address != "googlealerts-noreply@google.com" {
				continue
			}
		}
		// Process each message's part
		for {
			p, err := mr.NextPart()
			if err == io.EOF {
				break
			} else if err != nil {
				log.Fatal(err)
			}

			switch p.Header.(type) {
			case *mail.InlineHeader:
				// This is the message's text (can be plain-text or HTML)
				b, _ := ioutil.ReadAll(p.Body)
				content := string(b)
				scanner := bufio.NewScanner(strings.NewReader(content))

				pauta, veiculo, link := "", "", ""
				idx := -1
				newLine := false

				for scanner.Scan() {
					if newLine {
						if idx == 0 {
							pauta = scanner.Text()
						}
						if idx == 1 {
							veiculo = scanner.Text()
						}
						if strings.Contains(scanner.Text(), "https://") {
							stp := strings.Split(scanner.Text(), "url=")
							if len(stp) > 1 {
								link = strings.Split(scanner.Text(), "url=")[1]
								link = strings.Replace(link, ">", "", 1)
								link = strings.Split(link, "&ct=ga")[0]
							} else {
								link = scanner.Text()
							}
						}
						idx++
						if pauta != "" && veiculo != "" && link != "" {
							file.AddRow("Planilha1", []string{date.Format("2006-01-02 15:04:05"), pauta, veiculo, link})
							pauta, veiculo, link = "", "", ""
							newLine = false
							idx = -1
						}
					}
					if scanner.Text() == "" {
						newLine = true
						idx = 0
					}
					if strings.Contains(scanner.Text(), "- - - - - - - - - - - - - - - - - - - -") {
						break
					}
				}

				break
			}
		}

	}

	if err := <-done; err != nil {
		log.Fatal(err)
	}

	var fileName string

	fmt.Print("Digite o nome do arquivo (sem .xlsx): ")

	fmt.Scanf("%v", &fileName)

	log.Println("Salvando arquivo!")
	file.Save(fileName + ".xlsx")

	log.Println("Finalizado!")
	time.Sleep(time.Second * 2)
}
