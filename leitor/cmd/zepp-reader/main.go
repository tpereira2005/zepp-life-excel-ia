// Read-only Zepp export. Secrets travel over stdin/stdout, never command arguments or files.
package main

import (
 "encoding/json"
 "fmt"
 "os"
 "github.com/AlexxIT/SmartScaleConnect/pkg/zepp"
)

func main() {
 var request struct { Token, Account, Password, Code string }
 if err := json.NewDecoder(os.Stdin).Decode(&request); err != nil { fail("Pedido invalido") }
 client := zepp.NewClient()
 var err error
 switch {
 case request.Token != "": err = client.LoginWithToken(request.Token)
 case request.Code != "": err = client.LoginWithCode(request.Code)
 case request.Account != "" && request.Password != "": err = client.Login(request.Account, request.Password)
 default: fail("Autenticacao em falta")
 }
 if err != nil { fail(err.Error()) }
 weights, err := client.GetAllWeights()
 if err != nil { fail(err.Error()) }
 if len(weights) == 0 { fail("A Zepp nao devolveu medicoes; a exportacao anterior foi preservada") }
 if err := json.NewEncoder(os.Stdout).Encode(map[string]interface{}{"token":client.Token(), "weights":weights}); err != nil { fail("Erro ao exportar") }
}
func fail(message string) { fmt.Fprintln(os.Stderr, message); os.Exit(1) }
