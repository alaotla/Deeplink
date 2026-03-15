Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    const analyzeBtn = document.getElementById("analyze");
    const resultDiv = document.getElementById("result");
    const loading = document.getElementById("loading");
    const actionBar = document.getElementById("action-bar");
    const modeSelect = document.getElementById("mode-select");

    if (!analyzeBtn) return;

    analyzeBtn.onclick = async () => {
      // 增加安全检查，防止读取 null 的 value
      const mode = modeSelect ? modeSelect.value : "academic";
      
      resultDiv.innerText = "";
      loading.classList.remove("hidden");
      actionBar.classList.add("hidden");

      Office.context.mailbox.item.body.getAsync("text", async (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          const emailBody = res.value;
          const prompt = mode === "academic" 
            ? "你是一个专业的学术邮件助手。请用中文分析这封邮件的意图、专业术语，并给出礼貌得体的学术回复建议。使用Markdown格式。" 
            : "请用中文极简总结这封邮件的3个要点。";

          try {
            // 务必确保这里的 API Key 字符串不包含任何非 ASCII 字符或空格
            const apiKey = "sk-af54288ac4a040578283724defc90af4".trim(); 
            
            const response = await fetch("https://api.deepseek.com/chat/completions", {
              method: "POST",
              headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + apiKey
              },
              body: JSON.stringify({
                model: "deepseek-chat",
                messages: [{ role: "system", content: prompt }, { role: "user", content: emailBody }]
              })
            });

            if (!response.ok) throw new Error("API响应错误");

            const data = await response.json();
            const content = data.choices[0].message.content;

            loading.classList.add("hidden");
            actionBar.classList.remove("hidden");
            
            if (window.renderMarkdown) {
                window.renderMarkdown(content);
            } else {
                resultDiv.innerText = content;
            }

            document.getElementById("copy-btn").onclick = () => {
              navigator.clipboard.writeText(content);
              alert("已复制！");
            };
          } catch (e) {
            loading.classList.add("hidden");
            resultDiv.innerText = "分析失败: " + e.message;
          }
        }
      });
    };
  }
});