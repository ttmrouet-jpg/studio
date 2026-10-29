/**
 * Suivi d'utilisation du Studio Vidéo Square Habitat
 * ---------------------------------------------------
 * 1. Créer une Google Sheet vide (ex. « Suivi Studio Vidéo »).
 * 2. Menu Extensions → Apps Script, coller ce code, enregistrer.
 * 3. Déployer → Nouveau déploiement → type « Application Web »
 *    - Exécuter en tant que : Moi
 *    - Qui a accès : Tout le monde
 * 4. Copier l'URL de l'application Web (…/exec) et la transmettre à Claude
 *    (elle sera collée dans la constante STATS_URL du studio).
 */
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Vidéos') || ss.insertSheet('Vidéos');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Date', 'Agence', 'Négociateur', 'Modèle', 'Format', 'Durée (s)', 'Ville', 'Prix']);
    sh.setFrozenRows(1);
  }
  var d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) {}
  sh.appendRow([new Date(d.date || new Date()), d.agence || '', d.negociateur || '', d.modele || '', d.format || '',
                d.duree || '', d.ville || '', d.prix || '']);
  ensureSummary_(ss);
  return ContentService.createTextOutput('ok');
}

function ensureSummary_(ss) {
  if (ss.getSheetByName('Synthèse')) return;
  var s = ss.insertSheet('Synthèse');
  s.getRange('A1').setValue('Vidéos par agence et par mois');
  s.getRange('A3').setFormula('=QUERY(\'Vidéos\'!A:H,"select B, count(A) where A is not null group by B pivot year(A)*100+month(A)+1 label B \'Agence\'",1)');
  s.getRange('A15').setValue('Vidéos par négociateur (30 derniers jours)');
  s.getRange('A17').setFormula('=QUERY(\'Vidéos\'!A:H,"select C, B, count(A) where A > date \'"&TEXT(TODAY()-30,"yyyy-mm-dd")&"\' group by C, B order by count(A) desc label C \'Négociateur\', B \'Agence\', count(A) \'Vidéos\'",1)');
}
