package zepp

import (
 "io"
 "net/http"
 "strings"
 "testing"
)

type fakeTransport func(*http.Request)(*http.Response,error)
func(f fakeTransport) RoundTrip(r *http.Request)(*http.Response,error){return f(r)}
func TestExportPrimaryIdentity(t *testing.T){
 c:=NewClient()
 c.client.Transport=fakeTransport(func(r *http.Request)(*http.Response,error){
  if !strings.Contains(r.URL.Path,"/members/-1/weightRecords"){t.Fatal("not primary member")}
  return &http.Response{StatusCode:200,Body:io.NopCloser(strings.NewReader(`{"items":[{"weightType":0,"generatedTime":1767600000,"summary":{"weight":80}}],"next":0}`)),Header:make(http.Header)},nil
 })
 rows,err:=c.GetAllWeights()
 if err!=nil||len(rows)!=1||rows[0].ProfileID==nil||*rows[0].ProfileID != -1 {t.Fatalf("invalid export: %v",err)}
}
func TestRejectHTTPFailure(t *testing.T){
 c:=NewClient()
 c.client.Transport=fakeTransport(func(r *http.Request)(*http.Response,error){return &http.Response{StatusCode:401,Body:io.NopCloser(strings.NewReader(`{}`)),Header:make(http.Header)},nil})
 if _,err:=c.GetAllWeights();err==nil{t.Fatal("HTTP failure must not become an empty success")}
}
