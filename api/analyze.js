const axios = require('axios');

export default async function handler(req, res) {
  // 1. 设置跨域头（必须设置，否则插件无法访问这个接口）
  res.setHeader('Access-Control-Allow-Credentials', true);
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

  // 处理预检请求
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  // 只允许 POST 请求
  if (req.method !== 'POST') {
    return res.status(405).json({ error: '仅支持 POST 请求' });
  }

  const { emailBody, mode } = req.body;
  
  // 2. 读取 Vercel 环境变量里的 Key（这行不需要改）
  const DEEPSEEK_KEY = process.env.DEEPSEEK_KEY;

  if (!DEEPSEEK_KEY) {
    return res.status(500).json({ error: '服务器未配置 API Key' });
  }

  try {
    // 根据模式设置 Prompt
    const prompt = mode === "academic" 
      ? "你是一个顶尖的学术专家。请用中文分析这封邮件的意图、专业术语和行动要求，并给出得体且具有学术感的回复建议。请使用Markdown格式。" 
      : "请用中文极简总结这封邮件的3个要点。";

    // 3. 向 DeepSeek 发起请求
    const response = await axios.post('https://api.deepseek.com/chat/completions', {
      model: "deepseek-chat",
      messages: [
        { role: "system", content: prompt },
        { role: "user", content: emailBody }
      ],
      stream: false
    }, {
      headers: {
        'Authorization': `Bearer ${DEEPSEEK_KEY}`,
        'Content-Type': 'application/json'
      }
    });

    // 4. 将结果返回给插件
    res.status(200).json({ content: response.data.choices[0].message.content });

  } catch (error) {
    console.error('API Error:', error.response ? error.response.data : error.message);
    res.status(500).json({ 
      error: 'AI 分析服务暂时不可用',
      details: error.response ? error.response.data : error.message 
    });
  }
}
