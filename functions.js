/* global Office */
Office.initialize = () => {};

function open(url) {
  Office.context.ui.openBrowserWindow(url);
}

function encode(str) {
  return encodeURIComponent(str || "").replace(/%20/g, "+");
}

function normalizeName(details) {
  if (details && details.displayName) return details.displayName;
  if (details && details.emailAddress) {
    const local = details.emailAddress.split("@")[0].replace(/[._\-]+/g, " ");
    return local;
  }
  return "";
}

function openLinkedIn(event) {
  try {
    const item = Office.context.mailbox.item;
    if (!item) return event.completed();

    // Cible expéditeur (réception) ou premier destinataire (envoyé)
    let target = item.from || (item.to && item.to.length ? item.to[0] : null);
    const displayName = normalizeName(target);

    // Recherche LinkedIn basée sur le nom/prénom
    const name = displayName || (item.subject || "");
    open("https://www.linkedin.com/search/results/people/?keywords=" + encode(name));
  } finally {
    event.completed();
  }
}

// Expose
window.openLinkedIn = openLinkedIn;
