import { getDashboardData, buildAIContext } from "./_data.js";

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const { question } = req.body || {};
  if (!question) return res.status(400).json({ error: "No question provided" });

  try {
    const data = await getDashboardData();
    const DATA_CONTEXT = buildAIContext(data);

    const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${process.env.GROQ_API_KEY}`
      },
      body: JSON.stringify({
        model: process.env.GROQ_MODEL || "llama-3.3-70b-versatile",
        max_tokens: 1024,
        messages: [
          { role: "system", content: DATA_CONTEXT },
          { role: "user", content: question }
        ]
      })
    });

    const llm = await response.json();
    if (!response.ok || llm.error) return res.status(500).json({ error: llm.error?.message || "AI provider error" });

    return res.status(200).json({
      answer: llm.choices?.[0]?.message?.content || "No response received.",
      dataSource: data.meta
    });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
