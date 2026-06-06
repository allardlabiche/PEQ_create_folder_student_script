/**
 * Crée un menu personnalisé dans Google Sheets lors de l'ouverture du fichier.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎓 Automatisation PEQ')
    .addItem('Générer les dossiers élèves', 'genererDossiersEleves')
    .addToUi();
}

/**
 * Fonction principale permettant de dupliquer le template, injecter les données
 * et filtrer les onglets selon l'option de chaque élève.
 */
function genererDossiersEleves() {
  // =========================================================================
  // CONFIGURATION : Remplacez les ID ci-dessous par ceux de vos fichiers
  // =========================================================================
  const TEMPLATE_ID = '1T6Y6_KU_sPD9n0md4MIo1-MhTruZKXDnTZrARKEUFpc';
  const LISTING_ID = '1YMs87I-S6VwxonXUjeI39EL8I7kAxlQtHPq6hkj9cDA';
  const REPARTITION_ID = '11icTLAFSl4QL-GPvtJBJ9049hNt868IwxOQhMk7D-tk';
  const CONFIG_CELLULE_ID = '1_VuaG2OsFAYpW29X8cClXhnvQvVH_5R1TN0QBt2jwCU';
  const DOSSIER_DESTINATION_ID = '163rsNwi2QAv0AdLPGcFC1NiD8oVmJoVc'; 
  // =========================================================================

  // 1. Chargement du fichier de configuration des cellules (Fichier 4)
  const configSs = SpreadsheetApp.openById(CONFIG_CELLULE_ID);
  const configSheet = configSs.getSheets()[0];
  const configData = configSheet.getDataRange().getValues();
  
  const cellMapping = {};
  for (let i = 0; i < configData.length; i++) {
    let cle = configData[i][0];
    let cellule = configData[i][1];
    if (cle && cellule) {
      // Normalisation en minuscules pour éviter les erreurs de casse
      cellMapping[cle.toString().trim().toLowerCase()] = cellule.toString().trim();
    }
  }

  // 2. Chargement de la répartition des onglets par option (Fichier 3)
  const repartSs = SpreadsheetApp.openById(REPARTITION_ID);
  const repartSheet = repartSs.getSheets()[0];
  const repartData = repartSheet.getDataRange().getValues();
  
  const optionsHeaders = repartData[0].map(h => h.toString().trim().toLowerCase());
  const ongletsParOption = {};
  
  // Construction de la liste des onglets à garder pour chaque option (colonne par colonne)
  optionsHeaders.forEach((option, colIndex) => {
    if (option) {
      ongletsParOption[option] = [];
      for (let rowIndex = 1; rowIndex < repartData.length; rowIndex++) {
        let nomOnglet = repartData[rowIndex][colIndex];
        if (nomOnglet && nomOnglet.toString().trim() !== "") {
          ongletsParOption[option].push(nomOnglet.toString().trim());
        }
      }
    }
  });

  // 3. Traitement du Listing des élèves (Fichier 2)
  const listingSs = SpreadsheetApp.openById(LISTING_ID);
  const listingSheet = listingSs.getSheets()[0];
  const listingData = listingSheet.getDataRange().getValues();
  
  const templateFile = DriveApp.getFileById(TEMPLATE_ID);
  const dossierDest = DriveApp.getFolderById(DOSSIER_DESTINATION_ID);

  let compteurSucces = 0;

  // Boucle sur chaque élève (on commence à la ligne 1 pour ignorer l'en-tête)
  for (let i = 1; i < listingData.length; i++) {
    let ligneEleve = listingData[i];
    
    // Si la ligne est vide (pas de nom ni prénom), on passe à la suivante
    if (!ligneEleve[0] && !ligneEleve[1]) continue;

    // Correspondance exacte avec l'ordre de vos 8 colonnes (Point 1)
    let donneesEleve = {
      "nom": ligneEleve[0],
      "prenom": ligneEleve[1],
      "naissance": ligneEleve[2],       // Fait le pont avec "Naissance" du fichier config
      "email tuteur 1": ligneEleve[3],
      "email tuteur 2": ligneEleve[4],  // Gère la 2ème colonne "Email Tuteur 1" du listing
      "langue": ligneEleve[5],
      "session": ligneEleve[6],
      "option": ligneEleve[7]
    };

    let nom = donneesEleve["nom"] ? donneesEleve["nom"].toString().trim() : "";
    let prenom = donneesEleve["prenom"] ? donneesEleve["prenom"].toString().trim() : "";
    let optionEleve = donneesEleve["option"] ? donneesEleve["option"].toString().trim().toLowerCase() : "";

    // SCRIPT - ACTION 1 : Créer un nouveau fichier à partir du template
    let nomNouveauFichier = `Dossier_PEQ_${nom}_${prenom}`;
    let copieFichier = templateFile.makeCopy(nomNouveauFichier, dossierDest);
    let nouveauSs = SpreadsheetApp.openById(copieFichier.getId());
    
    // Par défaut, l'encodage se fait sur le PREMIER onglet du template.
    let ongletCible = nouveauSs.getSheets()[0]; 

    // SCRIPT - ACTION 2 : Encoder les informations de l'élève (Listing -> Template)
    for (let cle in cellMapping) {
      if (donneesEleve[cle] !== undefined && donneesEleve[cle] !== "") {
        let celluleCible = cellMapping[cle];
        try {
          ongletCible.getRange(celluleCible).setValue(donneesEleve[cle]);
        } catch(e) {
          Logger.log(`⚠️ Erreur d'écriture pour ${nom} (Clé : ${cle}, Cellule : ${celluleCible}) : ${e.message}`);
        }
      }
    }

    // SCRIPT - ACTION 3 : Suppression des onglets selon l'option
    let ongletsAGarder = ongletsParOption[optionEleve] || [];
    
    if (ongletsAGarder.length > 0) {
      let tousLesOnglets = nouveauSs.getSheets();
      let ongletsASupprimer = [];
      
      // Identifier les onglets à supprimer
      tousLesOnglets.forEach(onglet => {
        let nomOnglet = onglet.getName().trim();
        if (!ongletsAGarder.includes(nomOnglet)) {
          ongletsASupprimer.push(onglet);
        }
      });

      // Sécurité Google Sheets : Il faut obligatoirement conserver au moins 1 onglet visible.
      if (ongletsASupprimer.length < tousLesOnglets.length) {
        ongletsASupprimer.forEach(onglet => {
          nouveauSs.deleteSheet(onglet);
        });
      } else {
        Logger.log(`⚠️ Alerte pour ${nom} ${prenom} : Aucun des onglets requis n'a été trouvé dans le template. Aucun onglet n'a été supprimé pour éviter un crash.`);
      }
    } else {
      Logger.log(`⚠️ Option '${optionEleve}' inconnue ou vide dans le fichier de répartition pour l'élève : ${nom} ${prenom}.`);
    }

    compteurSucces++;
    Logger.log(`✅ Fichier généré avec succès : ${nomNouveauFichier}`);
  }
  
  SpreadsheetApp.getUi().alert(`Opération terminée ! ${compteurSucces} dossiers élèves ont été créés.`);
}
