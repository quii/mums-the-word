package main

import (
	mumstheword "github.com/quii/mums-the-word/api"
	"net/http"
)

func main() {
	http.ListenAndServe(":8080", http.HandlerFunc(mumstheword.Handler))
}
