/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // La fonction sera appelée directement par le manifest
    console.log("Office.js chargé, add-in prêt");
  }
});

/**
 * Recherche LinkedIn avec le nom de l'expéditeur ou du contact
 * Cette fonction est appelée depuis le manifest via ExecuteFunction
 * IMPORTANT: Cette fonction doit être dans le scope global pour être accessible
 */
function searchLinkedIn(event) {
  const item = Office.context.mailbox.item;
  
  try {
    let fullName = null;
    
    // Vérifier le type d'élément
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      // Pour un message (mail), récupérer l'expéditeur
      const sender = item.from;
      if (sender) {
        fullName = sender.displayName || sender.emailAddress || null;
      }
    } 
    else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      // Pour un rendez-vous, récupérer l'organisateur
      const organizer = item.organizer;
      if (organizer) {
        fullName = organizer.displayName || organizer.emailAddress || null;
      }
      
      // Si pas d'organisateur, essayer les participants requis
      if (!fullName) {
        const requiredAttendees = item.requiredAttendees;
        if (requiredAttendees && requiredAttendees.length > 0) {
          fullName = requiredAttendees[0].displayName || requiredAttendees[0].emailAddress || null;
        }
      }
      
      // Si toujours pas de nom, essayer les participants optionnels
      if (!fullName) {
        const optionalAttendees = item.optionalAttendees;
        if (optionalAttendees && optionalAttendees.length > 0) {
          fullName = optionalAttendees[0].displayName || optionalAttendees[0].emailAddress || null;
        }
      }
    }
    
    if (fullName) {
      // Nettoyer le nom (enlever les adresses email si c'est juste un email)
      // Si c'est un email, on peut essayer d'extraire le nom de l'email
      if (fullName.includes('@') && !fullName.includes(' ')) {
        // C'est probablement juste un email, utiliser la partie avant @
        fullName = fullName.split('@')[0].replace(/[._]/g, ' ');
      }
      
      // Encoder le nom pour l'URL
      const encodedName = encodeURIComponent(fullName.trim());
      // Construire l'URL de recherche LinkedIn
      const linkedInUrl = `https://www.linkedin.com/search/results/people/?keywords=${encodedName}`;
      
      // Ouvrir LinkedIn dans le navigateur par défaut
      // Utiliser une technique qui fonctionne dans les add-ins Outlook
      try {
        // Méthode 1: Essayer d'ouvrir directement avec un élément <a> cliqué programmatiquement
        // Cette méthode fonctionne mieux que window.open dans les contextes d'add-ins
        const link = document.createElement('a');
        link.href = linkedInUrl;
        link.target = '_blank';
        link.rel = 'noopener noreferrer';
        link.style.display = 'none';
        document.body.appendChild(link);
        
        // Utiliser à la fois click() et dispatchEvent pour maximiser la compatibilité
        if (link.click) {
          link.click();
        } else {
          const clickEvent = new MouseEvent('click', {
            view: window,
            bubbles: true,
            cancelable: true
          });
          link.dispatchEvent(clickEvent);
        }
        
        // Retirer l'élément après un court délai
        setTimeout(() => {
          if (link.parentNode) {
            document.body.removeChild(link);
          }
        }, 100);
        
        console.log("Tentative d'ouverture de LinkedIn:", linkedInUrl);
      } catch (e) {
        console.error("Erreur avec la méthode directe:", e);
        // Fallback: Utiliser une page de redirection dans une boîte de dialogue
        try {
          const redirectUrl = `https://rise-4.github.io/Rise-OutlookAddins-Linkedin/redirect.html?url=${encodeURIComponent(linkedInUrl)}`;
          
          Office.context.ui.displayDialogAsync(
            redirectUrl,
            { 
              height: 60, 
              width: 50,
              displayInIframe: false
            },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Erreur lors de l'ouverture de la boîte de dialogue:", asyncResult.error);
                showErrorMessage("Impossible d'ouvrir LinkedIn. Veuillez copier cette URL: " + linkedInUrl);
              } else {
                // Écouter les messages de la boîte de dialogue pour la fermer
                const dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                  if (arg.message === 'close') {
                    dialog.close();
                  }
                });
              }
            }
          );
        } catch (e2) {
          console.error("Erreur lors de l'ouverture de LinkedIn:", e2);
          showErrorMessage("Erreur lors de l'ouverture de LinkedIn. URL: " + linkedInUrl);
        }
      }
    } else {
      showErrorMessage("Impossible de récupérer le nom. Veuillez sélectionner un mail ou un contact avec un expéditeur/organisateur.");
    }
  } catch (error) {
    console.error("Erreur lors de la recherche LinkedIn:", error);
    showErrorMessage("Une erreur est survenue: " + error.message);
  }
  
  // Indiquer que la commande est terminée
  event.completed();
}

/**
 * Affiche un message d'erreur dans une boîte de dialogue
 */
function showErrorMessage(message) {
  Office.context.ui.displayDialogAsync(
    'about:blank',
    { height: 30, width: 40 },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
          dialog.close();
        });
        const htmlContent = `
          <!DOCTYPE html>
          <html>
          <head>
            <meta charset="UTF-8">
            <style>
              body { 
                font-family: 'Segoe UI', Arial, sans-serif; 
                padding: 20px; 
                text-align: center; 
                background-color: #f5f5f5;
              }
              .message {
                background-color: white;
                padding: 20px;
                border-radius: 5px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                margin-bottom: 15px;
              }
              button {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 3px;
                cursor: pointer;
                font-size: 14px;
              }
              button:hover {
                background-color: #106ebe;
              }
            </style>
          </head>
          <body>
            <div class="message">
              <p>${message}</p>
            </div>
            <button onclick="window.close()">Fermer</button>
          </body>
          </html>
        `;
        dialog.messageChild(htmlContent);
      }
    }
  );
}

