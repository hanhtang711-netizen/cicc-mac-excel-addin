import { messageForError, messageForWarning } from "./app/userFeedback";

const messageCode = new URLSearchParams(window.location.search).get("code") ?? "excel_runtime_error";
const messageElement = document.querySelector<HTMLElement>("#feedback-message");

if (messageElement !== null) {
  messageElement.textContent = messageCode === "series_palette_reused"
    ? messageForWarning(messageCode)
    : messageForError({ code: messageCode });
}
