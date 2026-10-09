/**
 * Code.gs — Réception des scores du QCM INTERPOL DVI dans Google Sheets.
 *
 * MISE EN SERVICE (5 minutes) :
 *   1. Ouvrez le tableur « Scores QCM DVI » dans votre Drive.
 *   2. Menu « Extensions > Apps Script ». Un éditeur s'ouvre.
 *   3. Remplacez tout le contenu par le code de ce fichier, puis enregistrez.
 *   4. Cliquez sur « Déployer > Nouveau déploiement > type : Application Web » :
 *        - Exécuter en tant que : Moi (votre compte)
 *        - Accès : « Tous les utilisateurs possédant un compte Google »…
 *          ⚠️ Non : choisissez « Tout le monde » (les élèves n'envoient pas de requête authentifiée).
 *   5. Autorisez le script (Google affiche un avertissement standard : « Afficher la
 *      fenêtre non vérifiée > Aller sur… », c'est normal pour un script personnel).
 *   6. Copiez l'URL de déploiement (elle finit par /exec) et collez-la dans la
 *      constante SUBMIT_URL en tête de la page QCM (HTML).
 *
 * AFFICHAGE DES RÉSULTATS : ouvrez simplement le tableur. Chaque élève y ajoute
 * une ligne avec son identité, son score, son temps et le détail de ses réponses.
 *
 * MAJ DU SCRIPT : après modification du code, refaites « Déployer > Gérer les
 * déploiements > ✏️ > Version : Nouvelle version », pour que l'URL /exec prenne effet.
 */

var SHEET_NAME = 'Réponses';

function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000); // évite que deux envois simultanés se marchent dessus

    var data = JSON.parse(e.postData.contents);
    if (!data || !data.nom || !data.prenom || !data.service ||
        data.score === undefined || data.total === undefined) {
      return json_({ ok: false, error: 'donnees invalides' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // En-têtes à la première écriture
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Horodatage', 'Nom', 'Prénom', 'Service',
                       'Score', 'Total', 'Pourcentage', 'Temps (s)', 'Détail des réponses']);
      sheet.setFrozenRows(1);
      sheet.getRange('A1:I1').setFontWeight('bold');
    }

    // Détail lisible des réponses : "Q3 : donné B, correct A ; ..."
    var detail = (data.answers || []).map(function (a, i) {
      var lettres = ['A', 'B', 'C', 'D'];
      var donne = (a.given !== undefined && a.given !== null) ? lettres[a.given] : '—';
      var okk = a.given === a.correct_index ? '✔' : '✘';
      return 'Q' + (i + 1) + '(' + okk + ' ' + donne + '/' + lettres[a.correct_index] + ')';
    }).join(' ');

    sheet.appendRow([
      new Date(),
      String(data.nom).substring(0, 100),
      String(data.prenom).substring(0, 100),
      String(data.service).substring(0, 150),
      data.score, data.total, data.percentage, data.time_spent_sec || 0, detail
    ]);

    lock.releaseLock();
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// Facultatif : GET ?action=ping pour tester l'URL depuis un navigateur
function doGet() {
  return json_({ ok: true, message: 'Endpoint QCM DVI opérationnel.' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
