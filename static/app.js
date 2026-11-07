const log = document.querySelector("#log-lines");
if (log && window.EventSource) {
  const stream = new EventSource(log.dataset.stream);
  stream.onmessage = event => { log.textContent += event.data + "\n"; };
}
const status = document.querySelector("[data-status]");
if (status) {
  const refreshStatus = async () => {
    try {
      const response = await fetch("/api/status");
      const value = await response.json();
      const rows = status.querySelectorAll("tr");
      value.printers.forEach((printer, index) => {
        const row = rows[index + 1];
        if (row) {
          row.cells[4].textContent = printer.jobs_today;
          row.cells[5].textContent = printer.last_result;
        }
      });
    } catch (_) {}
  };
  window.setInterval(refreshStatus, 5000);
  document.addEventListener("printerone-refresh", refreshStatus);
}
document.querySelectorAll('form[method="post"]').forEach(form => {
  form.addEventListener("submit", async event => {
    if (!window.fetch) return;
    event.preventDefault();
    const response = await fetch(form.action, { method: "POST", body: new FormData(form) });
    const message = document.querySelector("[data-form-result]") || document.createElement("p");
    message.dataset.formResult = "";
    message.textContent = response.ok ? "Saved." : (await response.text());
    if (!message.parentNode) form.appendChild(message);
    if (response.ok) document.dispatchEvent(new CustomEvent("printerone-refresh"));
  });
});
