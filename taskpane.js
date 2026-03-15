Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    const analyzeBtn = document.getElementById("analyze");
    const resultDiv = document.getElementById("result");
    const loading = document.getElementById("loading");
    const modeSelect = document.getElementById("mode-select");

    analyzeBtn.onclick = async () => {
      resultDiv.innerText = "";
      loading.classList.remove("hidden");

      Office.context.mailbox.item.body.getAsync("text", async (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          try {
            // 这里填写你的 Vercel 后端地址
            const response = await fetch("https://deeplink-eosin.vercel.app/api/analyze", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                emailBody: res.value,
                mode: modeSelect.value
              })
            });

            if (!response.ok) {
              throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            loading.classList.add("hidden");
            
            if (data.content) {
              window.renderMarkdown(data.content);
            } else {
              resultDiv.innerText = "分析失败：" + (data.error || "未知错误");
            }
          } catch (e) {
            loading.classList.add("hidden");
            resultDiv.innerText = "连接服务器失败，请稍后再试。";
            console.error(e);
          }
        }
      });
    };
  }
});
