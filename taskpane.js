Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    const analyzeBtn = document.getElementById("analyze");
    const resultDiv = document.getElementById("result");
    const loading = document.getElementById("loading");
    const modeSelect = document.getElementById("mode-select");

    if (!analyzeBtn) return;

    analyzeBtn.onclick = async () => {
      resultDiv.innerText = "";
      loading.classList.remove("hidden");

      Office.context.mailbox.item.body.getAsync("text", async (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          try {
            // 【关键修改】把下面的链接换成你刚刚在 Vercel 看到的域名，记得保留末尾的 /api/analyze
            const response = await fetch("https://deeplink-eosin.vercel.app/api/analyze", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                emailBody: res.value,
                mode: modeSelect.value
              })
            });

            const data = await response.json();
            loading.classList.add("hidden");
            
            if (data.content) {
              window.renderMarkdown(data.content);
            } else {
              resultDiv.innerText = "分析失败：" + (data.error || "未知错误");
            }
          } catch (e) {
            loading.classList.add("hidden");
            resultDiv.innerText = "连接服务器失败，请检查网络";
          }
        }
      });
    };
  }
});
