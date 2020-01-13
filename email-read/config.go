package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
)

// Config struct
type Config struct {
	Imap     string `json:"imap"`
	Username string `json:"username"`
	Password string `json:"password"`
}

var config Config

func init() {
	b, err := ioutil.ReadFile("config.json")
	if err != nil {
		fmt.Print("Erro ao ler arquivo de configuração *", err)
	}
	err = json.Unmarshal(b, &config)
	if err != nil {
		fmt.Print("Erro ao converter arquivo de configuração *", err)
	}
}
