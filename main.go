package main

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"log"
	"mime"
	"os"
	"path/filepath"
	"strings"
	"time"

	//"net/textproto"
	"github.com/joho/godotenv"
	"github.com/sashabaranov/go-openai"
	"github.com/xuri/excelize/v2"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/gmail/v1"
)

const (
	excelPath       = "professors.xlsx"
	sheetName       = "UESTC"
	schoolName      = "School of Information and Software Engineering, University of Electronic Science and Technology of China (UESTC)"
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

	if err := generateDrafts(); err != nil {
		log.Fatalf("Error: %v", err)
	}

	log.Println("✅ Done! Check your Gmail Drafts folder.")
}

func generateDrafts() error {
	// Open Excel
	f, err := excelize.OpenFile(excelPath)
	if err != nil {
		return fmt.Errorf("open excel: %w", err)
	}
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("read rows: %w", err)
	}

	if len(rows) <= 1 {
		return fmt.Errorf("no data rows found")
	}

	// Load progress from file (to continue from where it left off)
	progressFile := "progress.txt"
	startIndex := 1 // skip header by default

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

		para, err := generateResearchParagraph(ctx, aiClient, research)
		if err != nil {
			log.Printf("OpenAI error for %s: %v (fallback used)", name, err)
			para = fmt.Sprintf("Your research on %s strongly aligns with my academic background and interests.", research)
		}

		subject, body := generateEmailBody(name, para)

		if err := createDraft(gSrv, email, subject, body, attachmentPath); err != nil {
			log.Printf("❌ Draft creation failed for %s: %v", name, err)
		} else {
			log.Printf("✅ Draft created for %s <%s>", name, email)
			count++
		}

		time.Sleep(2 * time.Second)
	}

	// Save new progress index
	newProgress := endIndex
	if newProgress > len(rows) {
		newProgress = len(rows)
	}
	os.WriteFile(progressFile, []byte(fmt.Sprintf("%d", newProgress)), 0644)

	log.Printf("Total drafts created: %d", count)
	log.Printf("Progress saved: next run will start from row %d", newProgress)
	return nil
}

func generateResearchParagraph(ctx context.Context, client *openai.Client, research string) (string, error) {
	prompt := fmt.Sprintf(`You are helping a student write a short email paragraph to a professor about their research.
The paragraph should sound human, polite, and written in simple, clear English — like a non-native speaker who writes carefully.

Guidelines:
- Write 2–3 sentences only.
- Avoid greetings or professor’s name.
- Keep grammar correct but simple (no complex vocabulary).
- Use phrases like “Your contribution to… was of great interest to me,” “I was interested in knowing how…,” “It is interesting to consider…”
- Avoid phrases like “I would love to learn more,” “I’m excited,” or “I am passionate.”
- Write in a natural, slightly formal tone — respectful but not robotic.
- Connect my background naturally to their work.

Professor’s research area: %s

My background: I have experience in backend development, Golang, distributed systems, and AI. 
Link this background naturally to their research.`, research)

	resp, err := client.CreateChatCompletion(ctx, openai.ChatCompletionRequest{
		Model: openAIModel,
		Messages: []openai.ChatCompletionMessage{
			{Role: "system", Content: "You are a professional academic writing assistant."},
			{Role: "user", Content: prompt},
		},
	})
	if err != nil {
		return "", err
	}

	if len(resp.Choices) == 0 {
		return "", fmt.Errorf("no choices returned")
	}

	return strings.TrimSpace(resp.Choices[0].Message.Content), nil
}

func generateEmailBody(profName, researchPara string) (string, string) {
	subject := "Request for Master's Supervision (September 2026 Intake)"

	body := fmt.Sprintf(`
<html>
<body>
Dear Professor %s,<br><br>

I hope you are in a good health. I am <b>Usama Shoukat</b>, and I recently completed my <b>Bachelor's in Computer Science</b> from the <b>Government College University of Faisalabad, Pakistan</b> with a CGPA of <b>3.51/4.00</b>. I have a strong background in <b>software engineering and backend systems</b> where I have worked with Golang, JavaScript, APIs, distributed systems and different AI tools.<br><br>

I want to apply for a <b>Master’s Program under your supervision at the %s</b> via the <b>Chinese Government Scholarship (CSC) 2026</b>.<br><br>

%s<br><br>

If you are accepting new students, I would like to contribute to your research and learn under your guidance for the <b>2026 intake</b>. I have attached my CV for your review.<br><br>

Looking forward to hearing back from you.<br><br>

Best regards,<br>
<b>Usama Shoukat</b><br>
WeChat ID: UsamaShoukatCS<br>
GitHub: <a href="https://github.com/usamashoukatcs">github.com/usamashoukatcs</a><br>
LinkedIn: <a href="https://www.linkedin.com/in/usama-shoukat/">linkedin.com/in/usama-shoukat</a>
</body>
</html>
`, profName, schoolName, researchPara)

	return subject, body
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

	// Email headers
	msg.WriteString(fmt.Sprintf("To: %s\r\n", to))
	msg.WriteString(fmt.Sprintf("Subject: %s\r\n", subject))
	msg.WriteString("MIME-Version: 1.0\r\n")
	msg.WriteString(fmt.Sprintf("Content-Type: multipart/mixed; boundary=%s\r\n", boundary))
	msg.WriteString("\r\n")

	// HTML part (replace line breaks with <br>)
	msg.WriteString(fmt.Sprintf("--%s\r\n", boundary))
	msg.WriteString("Content-Type: text/html; charset=utf-8\r\n\r\n")
	//htmlBody := "<html><body>" + strings.ReplaceAll(plainBody, "\n", "<br>") + "</body></html>\r\n"
	msg.WriteString(plainBody + "\r\n")

	// Attachment part
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

	// End boundary
	msg.WriteString(fmt.Sprintf("--%s--", boundary))

	// Encode the full email for Gmail API
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
	fmt.Printf("Go to the following URL in your browser and paste the authorization code:\n\n%s\n\n", authURL)
	fmt.Print("Enter code: ")
	var code string
	fmt.Scan(&code)
	tok, err := config.Exchange(context.Background(), strings.TrimSpace(code))
	if err != nil {
		log.Fatalf("Token exchange error: %v", err)
	}
	return tok
}
