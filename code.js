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
 * Fonction principale permettant de dupliquer le template, filtrer les onglets,
 * puis injecter les données exclusivement dans l'onglet "Config".
 */
function genererDossiersEleves() {
  // Verrou pour éviter les double-clics impromptus
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    SpreadsheetApp.getUi().alert("Le script est déjà en cours d'exécution. Veuillez patienter.");
    return;
  }

  try {
    const ui = SpreadsheetApp.getUi();

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
        cellMapping[cle.toString().trim().toLowerCase()] = cellule.toString().trim();
      }
    }

    // 2. Chargement de la répartition des onglets par option (Fichier 3)
    const repartSs = SpreadsheetApp.openById(REPARTITION_ID);
    const repartSheet = repartSs.getSheets()[0];
    const repartData = repartSheet.getDataRange().getValues();
    
    const optionsHeaders = repartData[0].map(h => h.toString().trim().toLowerCase());
    const ongletsParOption = {};
    
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
    const dossierRacineDest = DriveApp.getFolderById(DOSSIER_DESTINATION_ID);

    let compteurSucces = 0;
    let compteurDoublonsIgnores = 0;
    let compteurEcrases = 0;

    // Variable pour mémoriser le choix de l'utilisateur sur la gestion des doublons
    let choixDoublon = null;

    // Boucle sur chaque élève
    for (let i = 1; i < listingData.length; i++) {
      let ligneEleve = listingData[i];
      
      if (!ligneEleve[0] && !ligneEleve[1]) continue;

      let donneesEleve = {
        "nom": listingData[i][0],
        "prenom": listingData[i][1],
        "naissance": listingData[i][2],       
        "email tuteur 1": listingData[i][3],
        "email tuteur 2": listingData[i][4],  
        "langue": listingData[i][5],
        "session": listingData[i][6],
        "option": listingData[i][7],
        "4e en option": listingData[i][8],
        "Date Affichage Complet": listingData[i][9]
      };

      let nom = donneesEleve["nom"] ? donneesEleve["nom"].toString().trim() : "";
      let prenom = donneesEleve["prenom"] ? donneesEleve["prenom"].toString().trim() : "";
      let optionTexte = donneesEleve["option"] ? donneesEleve["option"].toString().trim() : "Sans Option";
      let optionEleveClé = optionTexte.toLowerCase();
      let sessionTexte = donneesEleve["session"] ? donneesEleve["session"].toString().trim() : "Sans Session";

      // Structure du nom de fichier "nom prenom - option"
      let nomNouveauFichier = `${nom} ${prenom} - ${optionTexte}`;

      // Création/Récupération de l'arborescence des dossiers
      let dossierPEQBiche = obtenirOuCreerDossier(dossierRacineDest, "PEQ Biche");
      let dossierSession = obtenirOuCreerDossier(dossierPEQBiche, sessionTexte);
      let dossierOption = obtenirOuCreerDossier(dossierSession, optionTexte);

      // VERIFICATION DOUBLON
      let fichiersExistants = dossierOption.getFilesByName(nomNouveauFichier);
      let fichierCible = null;
      let existeDeja = false;

      if (fichiersExistants.hasNext()) {
        existeDeja = true;
        fichierCible = fichiersExistants.next();
        
        // Si c'est le premier doublon croisé, on demande une bonne fois pour toutes à l'utilisateur quoi faire
        if (choixDoublon === null) {
          let reponse = ui.alert(
            '⚠️ Fichiers existants détectés',
            'Certains dossiers élèves existent déjà dans les dossiers d\'options.\n\n' +
            'Voulez-vous ÉCRASER leurs informations dans l\'onglet Config (Oui) ou bien les IGNORER (Non) ?',
            ui.ButtonSet.YES_NO_CANCEL
          );
          
          if (reponse === ui.Button.YES) {
            choixDoublon = "ecraser";
          } else if (reponse === ui.Button.NO) {
            choixDoublon = "ignorer";
          } else {
            Logger.log("❌ Opération annulée par l'utilisateur.");
            return; // Arrêt complet du script
          }
        }
      }

      let nouveauSs;

      // TRAITEMENT SELON LE CHOIX DES DOUBLONS
      if (existeDeja) {
        if (choixDoublon === "ignorer") {
          Logger.log(`⏭️ Doublon ignoré pour l'élève : ${nomNouveauFichier}`);
          compteurDoublonsIgnores++;
          continue; // On passe à l'élève suivant
        } else if (choixDoublon === "ecraser") {
          Logger.log(`🔄 Mode Écrasement pour l'élève : ${nomNouveauFichier}`);
          nouveauSs = SpreadsheetApp.openById(fichierCible.getId());
          compteurEcrases++;
        }
      } else {
        // CAS NORMAL : Le fichier n'existe pas, on crée une copie du template
        let copieFichier = templateFile.makeCopy(nomNouveauFichier, dossierOption);
        nouveauSs = SpreadsheetApp.openById(copieFichier.getId());
        
        // SCRIPT - ACTION 2 : Suppression des onglets selon l'option (Uniquement sur les NOUVEAUX fichiers)
        let ongletsAGarder = ongletsParOption[optionEleveClé] || [];
        if (ongletsAGarder.length > 0) {
          let tousLesOnglets = nouveauSs.getSheets();
          let ongletsASupprimer = [];
          
          tousLesOnglets.forEach(onglet => {
            let nomOnglet = onglet.getName().trim();
            if (!ongletsAGarder.includes(nomOnglet) && nomOnglet.toLowerCase() !== "config") {
              ongletsASupprimer.push(onglet);
            }
          });

          if (ongletsASupprimer.length < tousLesOnglets.length) {
            ongletsASupprimer.forEach(onglet => {
              nouveauSs.deleteSheet(onglet);
            });
          }
        }
        compteurSucces++;
      }

      // SCRIPT - ACTION 3 : Dans les deux cas (Nouveau OU Écrasé), on écrit les données dans l'onglet "Config"
      let ongletConfig = nouveauSs.getSheetByName("Config");
      
      if (ongletConfig) {
        for (let cle in cellMapping) {
          if (donneesEleve[cle] !== undefined && donneesEleve[cle] !== "") {
            let celluleCible = cellMapping[cle];
            try {
              ongletConfig.getRange(celluleCible).setValue(donneesEleve[cle]);
            } catch(e) {
              Logger.log(`⚠️ Erreur d'écriture dans l'onglet [Config] pour ${nom} : ${e.message}`);
            }
          }
        }
      } else {
        Logger.log(`❌ Erreur critique : L'onglet "Config" est introuvable dans le fichier de ${nomNouveauFichier}.`);
      }

      // Force la sauvegarde immédiate dans le Drive
      SpreadsheetApp.flush();
    }
    
    // Message de fin détaillé
    let messageResultat = `Opération terminée !\n\n📊 Résultat :\n- ${compteurSucces} nouveau(x) dossier(s) créé(s).`;
    if (compteurEcrases > 0) {
      messageResultat += `\n- ${compteurEcrases} dossier(s) existant(s) mis à jour (données écrasées).`;
    }
    if (compteurDoublonsIgnores > 0) {
      messageResultat += `\n- ${compteurDoublonsIgnores} élève(s) ignoré(s) (fichiers déjà existants).`;
    }
    ui.alert(messageResultat);

  } catch (erreur) {
    SpreadsheetApp.getUi().alert("Une erreur est survenue : " + erreur.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Fonction utilitaire qui cherche un sous-dossier par son nom.
 * S'il n'existe pas, elle le crée automatiquement dans le dossier parent fourni.
 */
function obtenirOuCreerDossier(dossierParent, nomDossier) {
  const dossiersTrouves = dossierParent.getFoldersByName(nomDossier);
  if (dossiersTrouves.hasNext()) {
    return dossiersTrouves.next();
  }
  return dossierParent.createFolder(nomDossier);
}
