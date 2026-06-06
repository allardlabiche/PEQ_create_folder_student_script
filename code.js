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
    const dossierDest = DriveApp.getFolderById(DOSSIER_DESTINATION_ID);

    let compteurSucces = 0;
    let compteurDoublons = 0;

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
        "option": listingData[i][7]
      };

      let nom = donneesEleve["nom"] ? donneesEleve["nom"].toString().trim() : "";
      let prenom = donneesEleve["prenom"] ? donneesEleve["prenom"].toString().trim() : "";
      let optionEleve = donneesEleve["option"] ? donneesEleve["option"].toString().trim().toLowerCase() : "";

      let nomNouveauFichier = `Dossier_PEQ_${nom}_${prenom}`;

      // ANTI-DOUBLON : Vérification si le fichier existe déjà
      let fichiersExistants = dossierDest.getFilesByName(nomNouveauFichier);
      if (fichiersExistants.hasNext()) {
        Logger.log(`⏭️ Le fichier existe déjà pour l'élève ${nom} ${prenom}. Ligne ignorée.`);
        compteurDoublons++;
        continue; 
      }

      // SCRIPT - ACTION 1 : Créer un nouveau fichier à partir du template
      let copieFichier = templateFile.makeCopy(nomNouveauFichier, dossierDest);
      let nouveauSs = SpreadsheetApp.openById(copieFichier.getId());
      
      // SCRIPT - ACTION 2 : Suppression des onglets selon l'option
      let ongletsAGarder = ongletsParOption[optionEleve] || [];
      
      if (ongletsAGarder.length > 0) {
        let tousLesOnglets = nouveauSs.getSheets();
        let ongletsASupprimer = [];
        
        tousLesOnglets.forEach(onglet => {
          let nomOnglet = onglet.getName().trim();
          // SÉCURITÉ : On ne supprime JAMAIS l'onglet nommé "Config" (insensible à la casse)
          if (!ongletsAGarder.includes(nomOnglet) && nomOnglet.toLowerCase() !== "config") {
            ongletsASupprimer.push(onglet);
          }
        });

        if (ongletsASupprimer.length < tousLesOnglets.length) {
          ongletsASupprimer.forEach(onglet => {
            nouveauSs.deleteSheet(onglet);
          });
        } else {
          Logger.log(`⚠️ Alerte pour ${nom} ${prenom} : Aucun des onglets requis n'a été trouvé dans le template.`);
        }
      } else {
        Logger.log(`⚠️ Option '${optionEleve}' inconnue ou vide pour l'élève : ${nom} ${prenom}.`);
      }

      // SCRIPT - ACTION 3 : Encoder les informations de l'élève UNIQUEMENT dans l'onglet "Config"
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
        Logger.log(`❌ Erreur critique pour ${nom} ${prenom} : L'onglet nommé "Config" est introuvable dans ce template.`);
      }

      // Force l'enregistrement immédiat dans Google Drive
      SpreadsheetApp.flush();

      compteurSucces++;
      Logger.log(`✅ Fichier généré et complété avec succès : ${nomNouveauFichier}`);
    }
    
    // Message de fin
    let messageResultat = `Opération terminée !\n\n📊 Résultat :\n- ${compteurSucces} dossier(s) créé(s) et complété(s).`;
    if (compteurDoublons > 0) {
      messageResultat += `\n- ${compteurDoublons} élève(s) ignoré(s) car leur fichier existait déjà.`;
    }
    SpreadsheetApp.getUi().alert(messageResultat);

  } catch (erreur) {
    SpreadsheetApp.getUi().alert("Une erreur est survenue : " + erreur.message);
  } finally {
    lock.releaseLock();
  }
}
