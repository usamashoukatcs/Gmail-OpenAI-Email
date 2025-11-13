package followup

import (
	"context"
	"fmt"
	"strings"
)

func generateFollowupEmailBody(professorName, universityName, topics string) string {
	var topicLine string
	if topics != "" {
		topicLine = fmt.Sprintf("<p>I’m particularly interested in your research on <strong>%s</strong>, and I believe my background in backend development and Golang aligns well with these areas.</p>", topics)
	} else {
		topicLine = "<p>I remain deeply interested in your ongoing research directions.</p>"
	}

	return fmt.Sprintf(`
<html>
<body>
<p>Dear Professor %s,</p>

<p>I hope you are doing well. I’m writing to follow up on my previous message to express my continued interest in joining your research group at <strong>%s</strong>.</p>

%s

<p>I’m eager to contribute meaningfully and willing to put in the work required to progress effectively under your guidance.</p>

<p>For your convenience, I have reattached my CV.</p>

<p>Thank you for your time and kind consideration.</p>

<p>Best regards,<br>
<strong>Usama Shoukat</strong><br>
WeChat ID: UsamaShoukatCS<br>
GitHub: <a href="https://github.com/usamashoukatcs">github.com/usamashoukatcs</a><br>
LinkedIn: <a href="https://www.linkedin.com/in/usama-shoukat/">linkedin.com/in/usama-shoukat</a>
</p>
</body>
</html>
`, professorName, universityName, topicLine)

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

My background: I have experience in backend development, Golang, software engineering, distributed systems, and AI. 
Link this background naturally to their research.`, research)

	resp, err := client.CreateChatCompletion(ctx, openai.ChatCompletionRequest{
		Model: openAIModel,
		Messages: []openai.ChatCompletionMessage{
			{Role: "system", Content: "You are an expert academic assistant who only extracts key research topics from text."},
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
	<html> <body> Dear Professor %s,<br><br> 
	I hope you are in a good health. I am <b>Usama Shoukat</b>, and I recently completed my <b>Bachelor's in Computer Science</b> 
from the <b>Government College University of Faisalabad, Pakistan</b> with a CGPA of <b>3.51/4.00</b>. 
I have a strong background in <b>software engineering and backend systems</b> where I have worked with Golang, JavaScript, 
APIs, distributed systems and different AI tools.<br><br> 
	I want to apply for a <b>Master’s Program under your supervision at the %s</b> via the <b>%s 2026</b>.<br><br> 
	%s
	<br><br> 
	If you are accepting new students, I would be honored contribute to your research and learn under your guidance for the 
<b>2026 intake</b>. I have attached my CV for your review.<br><br> 
	Looking forward to hearing back from you.<br><br> Best regards,<br> 
	<b>Usama Shoukat</b><br> 
	WeChat ID: UsamaShoukatCS<br>
	</body> 
	</html>
	`, profName, schoolName, researchPara, scholarshipType)

	return subject, body
}
