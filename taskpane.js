function sendMail() {
  const value = document.getElementById("inputValue").value;
  let message = "";

  if (value === "1") message = "Selam";
  else if (value === "2") message = "Nasılsın";
  else if (value === "3") message = "Merhaba";
  else {
    alert("Geçerli bir değer girin (1, 2 veya 3)");
    return;
  }

  Office.context.mailbox.item.to.setAsync([{ emailAddress: "edegirmencifb@gmail.com" }]);
  Office.context.mailbox.item.subject.setAsync("Deneme Mail");
  Office.context.mailbox.item.body.setAsync(message, { coercionType: "html" });
}