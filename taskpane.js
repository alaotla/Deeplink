Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    const analyzeBtn = document.getElementById("analyze");
    const resultDiv = document.getElementById("result");
    const loading = document.getElementById("loading");
    const actionBar = document.getElementById("action-bar");
    const modeSelect = document.getElementById("mode-select");
    const keyInput = document.getElementById("api-key-input");

    // 尝试从本地缓存读取 Key，方便下次使用
    const savedKey = localStorage.getItem("deeplink_key");
    if (savedKey && keyInput) {
        keyInput.value = savedKey;
    }

    if (!analyzeBtn) return;

    analyzeBtn.onclick = async () => {
      const apiKey = keyInput ? keyInput.value.trim() : "";
      const mode = modeSelect ? modeSelect.value : "academic";

      if (!apiKey) {
        alert("请输入您的 DeepSeek API Key 才能开始分析！");
        return;
      }

      // 保存 Key 到本地缓存
      localStorage.setItem("deeplink_key", apiKey);
      
      resultDiv.innerText = "";
      loading.classList.remove("hidden");
      actionBar.classList.add("hidden");

      Office.context.mailbox.item.body.getAsync("text", async (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          const prompt = mode === "academic" 
            ? "你是一个顶尖的学术翻译专家。请用中文分析这封邮件的深度意图、专业术语和行动要求，并给出得体且具有学术感的回复建议。请使用Markdown格式。" 
            : "请用中文极简总结这封邮件的 3 个核心要点。";

          try {
            const response = await fetch("https://api.deepseek.com/chat/completions", {
              method: "POST",
              headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + apiKey
              },
              body: JSON.stringify({
                model: "deepseek-chat",
                messages: [
                  { role: "system", content: prompt },
                  { role: "user", content: res.value }
                ]
              })
            });

            if (!response.ok) throw new Error("API 响应错误，请检查 Key 是否有效。");

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
              alert("分析结果已复制！");
            };
          } catch (e) {
            loading.classList.add("hidden");
            resultDiv.innerText = "分析失败：" + e.message;
            console.error(e);
          }
        }
      });
    };
  }
});
