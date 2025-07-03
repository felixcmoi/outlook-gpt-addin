Office.onReady(() => {
  const item = Office.context.mailbox.item;
  document.getElementById('emailSubject').innerText = item.subject;

  window.generateReply = async function () {
    const bulletPoints = document.getElementById("bulletPoints").value;
    const emailBody = await new Promise((resolve, reject) => {
      item.body.getAsync("text", result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject("Erreur lecture email");
        }
      });
    });

    const response = await fetch("https://ton-backend.vercel.app/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        email: emailBody,
        bullets: bulletPoints,
      }),
    });

    const data = await response.json();
    document.getElementById("output").innerText = data.reply;
  };
});

Office.onReady(() => {
  console.log("Office.js chargé !");
  document.body.insertAdjacentHTML("beforeend", "<p>Loaded depuis Outlook ✅</p>");
});
