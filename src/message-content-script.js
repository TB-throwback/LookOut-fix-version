browser.runtime.onMessage.addListener(message => {
  if (message.command != "showJunkTnefWarning") {
    return;
  }

  if (document.getElementById("lookout-junk-tnef-warning")) {
    return;
  }

  let banner = document.createElement("div");

  banner.id = "lookout-junk-tnef-warning";

  banner.textContent =
    "LookOut did not process winmail.dat because this message is in your Junk folder.";

  banner.style.cssText = `
    background: #fff3cd;
    color: #664d03;
    border: 1px solid #ffecb5;
    padding: 10px 14px;
    margin: 8px;
    font-family: sans-serif;
    font-weight: 600;
  `;

  document.body.prepend(banner);
});
