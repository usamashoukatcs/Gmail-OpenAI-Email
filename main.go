package main

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"log"
	"math/rand"
	"mime"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/joho/godotenv"
	"github.com/sashabaranov/go-openai"
	"github.com/xuri/excelize/v2"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/gmail/v1"
)

const (
	excelPath         = "professors.xlsx"
	scholarshipType   = "CSC Scholarship"
	followUpSheetName = "ZJU"
	//followUpSchoolName = "Zhejiang University"
	sheetName       = "UESTC"
	schoolName      = "University of Electronic Science and Technology of China (UESTC)"
	attachmentPath  = "UsamaShoukatCV.pdf"
	credentialsFile = "credentials.json"
	tokenFile       = "token.json"
	openAIModel     = "gpt-4o-mini"
	maxDrafts       = 30
)

func main() {
	log.Println("Starting email draft generator...")
	err := godotenv.Load()
	if err != nil {
		log.Fatal("Error loading .env file")
	}

	openaiKey := os.Getenv("OPENAI_API_KEY")
	if openaiKey == "" {
		log.Fatal("OPENAI_API_KEY not set")
	}

	mode := "initial"
	if len(os.Args) > 1 && (os.Args[1] == "followup" || os.Args[1] == "initial") {
		mode = os.Args[1]
	}

	log.Printf("Mode: %s", strings.ToUpper(mode))

	if err := generateDrafts(mode); err != nil {
		log.Fatalf("Error: %v", err)
	}

	log.Println("✅ Done! Check your Gmail Drafts folder.")
}

func generateDrafts(mode string) error {
	f, err := excelize.OpenFile(excelPath)
	if err != nil {
		return fmt.Errorf("open excel: %w", err)
	}

	sheet := sheetName
	if mode == "followup" {
		sheet = followUpSheetName
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return fmt.Errorf("read rows: %w", err)
	}

	if len(rows) <= 1 {
		return fmt.Errorf("no data rows found")
	}

	progressFile := sheet + "Progress.txt"
	if mode == "followup" {
		progressFile = "followUpProgress.txt"
	}

	startIndex := 1
	if data, err := os.ReadFile(progressFile); err == nil {
		var savedIndex int
		if _, err := fmt.Sscanf(string(data), "%d", &savedIndex); err == nil {
			startIndex = savedIndex
		}
	}

	endIndex := startIndex + maxDrafts
	log.Printf("Processing professors %d to %d ...", startIndex, endIndex)

	// Gmail + OpenAI setup
	gSrv, err := getGmailService()
	if err != nil {
		return fmt.Errorf("gmail auth: %w", err)
	}

	openAIKey := os.Getenv("OPENAI_API_KEY")
	if openAIKey == "" {
		return fmt.Errorf("OPENAI_API_KEY not set")
	}
	aiClient := openai.NewClient(openAIKey)
	ctx := context.Background()

	count := 0
	for i, row := range rows {
		if i < startIndex {
			continue
		}
		if i >= endIndex {
			break
		}

		if len(row) < 3 {
			continue
		}

		profInfo := strings.TrimSpace(row[1])
		research := strings.TrimSpace(row[2])
		if profInfo == "" || research == "" {
			continue
		}

		parts := strings.SplitN(profInfo, ":", 2)
		if len(parts) < 2 {
			continue
		}

		name := strings.TrimSpace(parts[0])
		email := strings.TrimSpace(parts[1])
		if name == "" || email == "" {
			continue
		}

		subject, body, err := generateEmail(ctx, aiClient, name, research, mode)
		if err != nil {
			log.Printf("❌ Email generation failed for %s: %v", name, err)
			continue
		}

		if err := createDraft(gSrv, email, subject, body, attachmentPath); err != nil {
			log.Printf("❌ Draft creation failed for %s: %v", name, err)
		} else {
			log.Printf("✅ Draft created for %s <%s>", name, email)
			count++
		}

		time.Sleep(2 * time.Second)
	}

	newProgress := endIndex
	if newProgress > len(rows) {
		newProgress = len(rows)
	}
	os.WriteFile(progressFile, []byte(fmt.Sprintf("%d", newProgress)), 0644)

	log.Printf("Total drafts created: %d", count)
	log.Printf("Progress saved: next run will start from row %d", newProgress)
	return nil
}

func generateEmail(ctx context.Context, client *openai.Client, profName, researchText, mode string) (string, string, error) {
	prompt := fmt.Sprintf(`
Extract 1–2 main research topics or directions from the professor’s research text below (keep it short and clear, no sentences).
Professor's research text:
%s
`, researchText)

	resp, err := client.CreateChatCompletion(ctx, openai.ChatCompletionRequest{
		Model: openAIModel,
		Messages: []openai.ChatCompletionMessage{
			{Role: "system", Content: "You are a concise academic assistant summarizing professors’ research fields in a few words."},
			{Role: "user", Content: prompt},
		},
		Temperature: 0.4,
	})
	if err != nil || len(resp.Choices) == 0 {
		log.Printf("⚠️ OpenAI topic extraction failed for %s: %v", profName, err)
		return "", "", err
	}

	researchTopics := strings.TrimSpace(resp.Choices[0].Message.Content)
	if researchTopics == "" {
		researchTopics = "computer science and related technologies"
	}

	researchPlan := `
My intended research plan involves exploring distributed systems and backend technologies using Go,
with an emphasis on building efficient, reliable, and scalable software systems.
I am also interested in integrating AI-based optimization or automation approaches into software engineering problems.
`

	intro := fmt.Sprintf(`
Respected Professor %s,<br><br>
I hope you are doing well.`, profName)

	if mode == "followup" {
		intro = fmt.Sprintf(`
Respected Professor %s,<br><br>
I hope you are doing well. I wanted to kindly follow up on my previous email regarding the possibility of pursuing a Master's degree under your supervision.`, profName)
	}

	subject := getRandomSubject()
	body := fmt.Sprintf(`
<html>
<body>
%s<br><br>

I am <b>Usama Shoukat</b> from Pakistan, and I have completed my Bachelor's in Computer Science
with a CGPA of <b>3.51/4.00</b> from Government College University Faisalabad, a well-known institution in Pakistan.<br><br>

I came across your research profile and was deeply impressed by your work in <b>%s</b>.
I find your research directions highly relevant to my academic background and interests.<br><br>

%s<br><br>

I am highly motivated to pursue a Master's degree under your supervision and intend to apply for the <b>%s</b>
(or any equivalent scholarship offered by your institution).<br><br>

If you find my profile suitable, it would be an honor to discuss the possibility of joining your research group.
I have attached my CV for your review.<br><br>

Thank you very much for your time and consideration.<br><br>

Best regards,<br>
<b>Usama Shoukat</b><br>
WeChat ID: UsamaShoukatCS<br>
</body>
</html>
`, intro, researchTopics, researchPlan, scholarshipType)

	return subject, body, nil
}

func getRandomSubject() string {
	subjects := []string{
		"Request for Master's Supervision (September 2026 Intake)",
		"Prospective Master's Student Interested in Your Research (2026 Intake)",
		"Supervision Inquiry for Master's Program (Fall 2026)",
		"Seeking Master's Supervision at Your Research Group (2026)",
		"Application for Master's Supervision - September 2026",
		"Interest in Joining Your Research Group for Master's 2026",
		"Inquiry Regarding Master's Supervision (2026 Admission)",
		"Exploring Master's Research Opportunities with You (2026)",
		"Request to Pursue Master's Studies Under Your Guidance (2026)",
		"Potential Master's Student Interested in Your Research Work",
	}
	r := rand.New(rand.NewSource(time.Now().UnixNano()))
	return subjects[r.Intn(len(subjects))]
}

func getGmailService() (*gmail.Service, error) {
	ctx := context.Background()
	b, err := os.ReadFile(credentialsFile)
	if err != nil {
		return nil, fmt.Errorf("read credentials.json: %w", err)
	}
	config, err := google.ConfigFromJSON(b, gmail.GmailComposeScope, gmail.GmailModifyScope)
	if err != nil {
		return nil, fmt.Errorf("config parse: %w", err)
	}

	tok, err := tokenFromFile(tokenFile)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(tokenFile, tok)
	}
	client := config.Client(ctx, tok)
	return gmail.New(client)
}

func createDraft(srv *gmail.Service, to, subject, body, attachmentFile string) error {
	rawMsg, err := buildRawEmail(to, subject, body, attachmentFile)
	if err != nil {
		return fmt.Errorf("build email: %w", err)
	}
	msg := &gmail.Message{Raw: rawMsg}
	_, err = srv.Users.Drafts.Create("me", &gmail.Draft{Message: msg}).Do()
	return err
}

func buildRawEmail(to, subject, plainBody, attachmentFile string) (string, error) {
	boundary := "BOUNDARY123"
	var msg strings.Builder

	msg.WriteString(fmt.Sprintf("To: %s\r\n", to))
	msg.WriteString(fmt.Sprintf("Subject: %s\r\n", subject))
	msg.WriteString("MIME-Version: 1.0\r\n")
	msg.WriteString(fmt.Sprintf("Content-Type: multipart/mixed; boundary=%s\r\n", boundary))
	msg.WriteString("\r\n")

	msg.WriteString(fmt.Sprintf("--%s\r\n", boundary))
	msg.WriteString("Content-Type: text/html; charset=utf-8\r\n\r\n")
	msg.WriteString(plainBody + "\r\n")

	data, err := os.ReadFile(attachmentFile)
	if err != nil {
		return "", err
	}
	_, fileName := filepath.Split(attachmentFile)
	mimeType := mime.TypeByExtension(filepath.Ext(fileName))
	if mimeType == "" {
		mimeType = "application/octet-stream"
	}

	msg.WriteString(fmt.Sprintf("--%s\r\n", boundary))
	msg.WriteString(fmt.Sprintf("Content-Type: %s; name=\"%s\"\r\n", mimeType, fileName))
	msg.WriteString("Content-Transfer-Encoding: base64\r\n")
	msg.WriteString(fmt.Sprintf("Content-Disposition: attachment; filename=\"%s\"\r\n\r\n", fileName))

	encoded := base64.StdEncoding.EncodeToString(data)
	for i := 0; i < len(encoded); i += 76 {
		end := i + 76
		if end > len(encoded) {
			end = len(encoded)
		}
		msg.WriteString(encoded[i:end] + "\r\n")
	}
	msg.WriteString(fmt.Sprintf("--%s--", boundary))

	return base64.URLEncoding.EncodeToString([]byte(msg.String())), nil
}

func tokenFromFile(path string) (*oauth2.Token, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}

func saveToken(path string, token *oauth2.Token) {
	f, err := os.Create(path)
	if err != nil {
		log.Fatalf("Unable to save oauth token: %v", err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(token)
	log.Printf("Token saved to %s", path)
}

func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline, oauth2.ApprovalForce)
	fmt.Printf("Go to the following URL and paste the authorization code:\n\n%s\n\n", authURL)
	fmt.Print("Enter code: ")
	var code string
	fmt.Scan(&code)
	tok, err := config.Exchange(context.Background(), strings.TrimSpace(code))
	if err != nil {
		log.Fatalf("Token exchange error: %v", err)
	}
	return tok
}
