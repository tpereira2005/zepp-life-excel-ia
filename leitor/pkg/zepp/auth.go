package zepp

import (
	"encoding/json"
	"fmt"
	"net/url"
	"strings"

	"github.com/AlexxIT/SmartScaleConnect/pkg/xiaomi"
	"github.com/google/uuid"
)

const paramsZeppLife = "_json=true&" +
	"client_id=428135909242707968&" +
	"pt=1&" +
	"redirect_uri=https://api-mifit-cn.huami.com/huami.health.loginview.do&" +
	"response_type=code"

func (c *Client) Login(username, password string) error {
	client := xiaomi.NewClient("")
	code, err := client.OAuth2(paramsZeppLife, username, password)
	if err != nil {
		return err
	}
	return c.LoginWithCode(code)
}

func (c *Client) LoginWithCode(code string) error {
	if code == "" {
		return fmt.Errorf("codigo OAuth Xiaomi vazio")
	}

	// country CN is OK
	form := url.Values{
		"app_name":     {"com.xiaomi.hm.health"},
		"app_version":  {"6.14.0"},
		"code":         {code},
		"country_code": {"CN"},
		"device_id":    {uuid.NewString()},
		"device_model": {"phone"},
		"dn":           {"api-mifit.zepp.com"},
		"grant_type":   {"request_token"},
		"third_name":   {"xiaomi-hm-mifit"},
	}

	res, err := c.client.Post(
		"https://account.zepp.com/v2/client/login", "application/x-www-form-urlencoded", strings.NewReader(form.Encode()),
	)
	if err != nil {
		return err
	}
	defer res.Body.Close()

	var res1 struct {
		TokenInfo struct {
			LoginToken string `json:"login_token"`
			AppToken   string `json:"app_token"`
			UserId     string `json:"user_id"`
			Ttl        int    `json:"ttl"`
			AppTtl     int    `json:"app_ttl"`
		} `json:"token_info"`
		RegistInfo struct {
			IsNewUser   int    `json:"is_new_user"`
			RegistDate  int64  `json:"regist_date"`
			Region      string `json:"region"`
			CountryCode string `json:"country_code"`
		} `json:"regist_info"`
		ThirdpartyInfo struct {
			Nickname     string `json:"nickname"`
			Icon         string `json:"icon"`
			ThirdId      string `json:"third_id"`
			AccessToken  string `json:"access_token"`
			RefreshToken string `json:"refresh_token"`
			ExpiresIn    int    `json:"expires_in"`
		} `json:"thirdparty_info"`
		Result string `json:"result"`
		Domain struct {
			IdDns string `json:"id-dns"`
		} `json:"domain"`
		Domains []struct {
			Cnames []string `json:"cnames"`
			Host   string   `json:"host"`
		} `json:"domains"`
	}

	if err = json.NewDecoder(res.Body).Decode(&res1); err != nil {
		return err
	}
	if res1.TokenInfo.AppToken == "" || res1.TokenInfo.UserId == "" {
		if res1.Result == "" {
			return fmt.Errorf("a Zepp rejeitou o codigo de autorizacao")
		}
		return fmt.Errorf("a Zepp rejeitou o codigo de autorizacao: %s", res1.Result)
	}

	c.appToken = res1.TokenInfo.AppToken
	c.userID = res1.TokenInfo.UserId

	return nil
}

type CodeClient struct {
	*Client
}

func NewCodeClient() *CodeClient {
	return &CodeClient{Client: NewClient()}
}

func (c *CodeClient) Login(_ string, code string) error {
	return c.Client.LoginWithCode(code)
}

func (c *Client) LoginWithToken(token string) error {
	c.userID, c.appToken, _ = strings.Cut(token, ":")
	return c.GetFamilyMembers()
}

func (c *Client) Token() string {
	return c.userID + ":" + c.appToken
}
