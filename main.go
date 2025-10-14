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
	sheetName       = "USTC"
	attachmentPath  = "UsamaShoukatCV.pdf"
	credentialsFile = "credentials.json"
	tokenFile       = "token.json"
	openAIModel     = "gpt-4o-mini"
	maxDrafts       = 25
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
	prompt := fmt.Sprintf(`Write a 2-4 sentence paragraph expressing genuine interest in a professor's research.
Keep it professional, natural, and relevant to backend engineering, Golang, distributed systems, and AI.

Research topic: %s

The paragraph should:
- Be of 2-4 sentences and in easy wordings.
- Show understanding of the research area.
- Connect it meaningfully with my academic background and goals.
- End with a thoughtful, natural sentence that reflects alignment or curiosity — not generic phrases like "I would love to learn more" or "I am excited to explore this field."`, research)

	resp, err := client.CreateChatCompletion(ctx, openai.ChatCompletionRequest{
		Model: openAIModel,
		Messages: []openai.ChatCompletionMessage{
			{Role: "system", Content: "You are an academic writing assistant."},
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
	subject := "Request for Master's Supervision"

	body := fmt.Sprintf(`Dear Professor %s,

I hope this message finds you well. My name is Usama Shoukat, and I recently completed my Bachelor’s in Computer Science (CGPA 3.51/4.00) from Government Graduate College of Science, Faisalabad (affiliated with GCUF), Pakistan. I am writing to express my interest in pursuing a Master’s degree under the ANSO Scholarship at the University of Science and Technology of China (USTC) under your supervision.

I have a strong background in software engineering, backend development, and intelligent systems, with practical experience in Golang-based scalable APIs, distributed systems, and AI integration. During my undergraduate studies, I completed several projects, including a real-time video and chat system using Golang and WebRTC, and an AI-powered Transcripto app (Text-to-Speech & Speech-to-Text) — both reflecting my passion for combining software development with applied intelligence and research.

%s

If you are currently considering graduate students for your team, I would be deeply honored to receive your guidance and supervision for the 2026 intake. My CV is attached for your kind review.

Thank you for your time and consideration. I look forward to hearing from you and working under your guidance.

Warm regards,
Usama Shoukat
WeChat ID: UsamaShoukatCS
GitHub: https://github.com/usamashoukatcs
LinkedIn: https://www.linkedin.com/in/usama-shoukat/
`, profName, researchPara)

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
	htmlBody := "<html><body>" + strings.ReplaceAll(plainBody, "\n", "<br>") + "</body></html>\r\n"
	msg.WriteString(htmlBody + "\r\n")

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
