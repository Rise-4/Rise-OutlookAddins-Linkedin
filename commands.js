/* global Office, posthog */

const APP_NAME = "outlook-linkedin-search";
const APP_VERSION = "1.1.0";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("Office.js chargé, add-in prêt");
    try {
      const email = Office.context.mailbox && Office.context.mailbox.userProfile
        ? Office.context.mailbox.userProfile.emailAddress
        : null;
      identifyUser(email);
      track("addin_loaded", {
        host_platform: (Office.context.diagnostics && Office.context.diagnostics.platform) || "unknown",
        host_version: (Office.context.diagnostics && Office.context.diagnostics.version) || "unknown"
      });
    } catch (e) {
      console.warn("Telemetry init failed:", e);
    }
  }
});

function identifyUser(email) {
  try {
    if (!email || typeof posthog === "undefined") return;
    const lower = email.toLowerCase();
    const tenant = lower.split("@")[1] || "unknown";
    posthog.identify(lower, { tenant });
    posthog.register({ app: APP_NAME, tenant, app_version: APP_VERSION });
  } catch (e) {}
}

function track(event, props) {
  try {
    if (typeof posthog !== "undefined" && posthog.capture) {
      posthog.capture(event, props || {}, { send_instantly: true });
    }
  } catch (e) {}
}

/**
 * Recherche LinkedIn avec le nom de l'expéditeur ou du contact
 * Cette fonction est appelée depuis le manifest via ExecuteFunction
 * IMPORTANT: Cette fonction doit être dans le scope global pour être accessible
 */
function searchLinkedIn(event) {
  const item = Office.context.mailbox.item;
  const source = detectSource(item);
  track("linkedin_search_clicked", { source });

  try {
    let fullName = null;

    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      const sender = item.from;
      if (sender) {
        fullName = sender.displayName || sender.emailAddress || null;
      }
    }
    else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      const organizer = item.organizer;
      if (organizer) {
        fullName = organizer.displayName || organizer.emailAddress || null;
      }

      if (!fullName) {
        const requiredAttendees = item.requiredAttendees;
        if (requiredAttendees && requiredAttendees.length > 0) {
          fullName = requiredAttendees[0].displayName || requiredAttendees[0].emailAddress || null;
        }
      }

      if (!fullName) {
        const optionalAttendees = item.optionalAttendees;
        if (optionalAttendees && optionalAttendees.length > 0) {
          fullName = optionalAttendees[0].displayName || optionalAttendees[0].emailAddress || null;
        }
      }
    }

    if (fullName) {
      let nameWasEmailFallback = false;
      if (fullName.includes('@') && !fullName.includes(' ')) {
        fullName = fullName.split('@')[0].replace(/[._]/g, ' ');
        nameWasEmailFallback = true;
      }

      const encodedName = encodeURIComponent(fullName.trim());
      const linkedInUrl = `https://www.linkedin.com/search/results/people/?keywords=${encodedName}`;

      openInDefaultBrowser(linkedInUrl, event, source, nameWasEmailFallback);
    } else {
      track("linkedin_search_failed", { source, reason: "no_name_available" });
      showNotification("Impossible de récupérer le nom. Veuillez sélectionner un mail ou un contact.");
      event.completed();
    }
  } catch (error) {
    console.error("Erreur lors de la recherche LinkedIn:", error);
    track("linkedin_search_error", { source, error_message: String(error && error.message || error) });
    showNotification("Une erreur est survenue: " + error.message);
    event.completed();
  }
}

/**
 * Identifie la surface d'où vient le clic (mail, organisateur de rdv, participant).
 */
function detectSource(item) {
  try {
    if (!item || !item.itemType) return "unknown";
    if (item.itemType === Office.MailboxEnums.ItemType.Message) return "message";
    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      return item.organizer ? "appointment_organizer" : "appointment_attendee";
    }
    return "unknown";
  } catch (e) {
    return "unknown";
  }
}

/**
 * Ouvre une URL dans le navigateur par défaut du système
 * Utilise openBrowserWindow (Mailbox 1.6+) avec fallback
 */
function openInDefaultBrowser(url, event, source, nameWasEmailFallback) {
  if (Office.context.ui.openBrowserWindow) {
    Office.context.ui.openBrowserWindow(url);
    console.log("LinkedIn ouvert dans le navigateur par défaut:", url);
    track("linkedin_search_succeeded", {
      source,
      opened_via: "openBrowserWindow",
      name_was_email_fallback: !!nameWasEmailFallback
    });
    event.completed();
  } else {
    console.warn("openBrowserWindow non disponible, utilisation du fallback");

    const redirectHtml = `https://rise-4.github.io/all--outlook-linkedin--addin/redirect.html?url=${encodeURIComponent(url)}`;

    Office.context.ui.displayDialogAsync(
      redirectHtml,
      { height: 10, width: 10, displayInIframe: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = asyncResult.value;
          track("linkedin_search_succeeded", {
            source,
            opened_via: "dialog_fallback",
            name_was_email_fallback: !!nameWasEmailFallback
          });
          setTimeout(() => {
            try {
              dialog.close();
            } catch (e) {
              // Dialogue peut déjà être fermé
            }
          }, 2000);
        } else {
          console.error("Erreur displayDialogAsync:", asyncResult.error);
          track("linkedin_search_error", {
            source,
            opened_via: "dialog_fallback",
            error_message: (asyncResult.error && asyncResult.error.message) || "displayDialogAsync_failed"
          });
          showNotification("Impossible d'ouvrir LinkedIn. URL: " + url);
        }
        event.completed();
      }
    );
  }
}

/**
 * Affiche une notification à l'utilisateur
 * Utilise l'API de notification d'Outlook si disponible
 */
function showNotification(message) {
  if (Office.context.mailbox.item && Office.context.mailbox.item.notificationMessages) {
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "linkedin-notification",
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: message,
        icon: "Icon.16x16",
        persistent: false
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Erreur notification:", result.error);
          console.log("Message pour l'utilisateur:", message);
        }
      }
    );
  } else {
    console.log("Message pour l'utilisateur:", message);
  }
}
