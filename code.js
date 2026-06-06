// =========================================================================
// CONFIGURATION GLOBALE
// =========================================================================
// Liste des adresses emails autorisées à TOUT voir et TOUT modifier (en minuscules)
const UTILISATEURS_AUTORISES = [
  "benjamin.allard@istlm.org",
  "peq-c3d-biche@istlm.org",
  "peq-c3d-mons@istlm.org"
];

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
 * Cette fonction DOIT être liée à un déclencheur installable "Lors de l'ouverture"
 * dans le fichier TEMPLATE (et s'appliquera sur les copies élèves).
 */
function verifierDateAffichageOngletsInstallable() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ongletConfig = ss.getSheetByName("Config");
    
    if (!ongletConfig) return; 

    // Récupération sécurisée de l'email de l'utilisateur qui ouvre le fichier
    const emailUtilisateurActuel = Session.getActiveUser().getEmail().toLowerCase().trim();

    // SI C'EST UN ADMINISTRATEUR : On lui affiche tout pour qu'il puisse travailler
    if (UTILISATEURS_AUTORISES.includes(emailUtilisateurActuel)) {
      const tousLesOnglets = ss.getSheets();
      tousLesOnglets.forEach(onglet => {
        if (onglet.getName() !== "Config") onglet.showSheet();
      });
      ongletConfig.showSheet(); 
      return; // On arrête ici, pas besoin de masquer pour l'admin
    }

    // --- LOGIQUE POUR LES ÉLÈVES (OU UTILISATEURS NON AUTORISÉS) ---
    
    // 1. Liste des onglets de matières soumis à la date d'affichage
    const ongletsAControler = [
      "Math-EQ", "FR-EQ", "FHG", "FSE-EQ", "FS-EQ", 
      "Rel", "EP", "LM1", "Math-P", "FR-P", "FSE-P", "FS-P"
    ];

    // 2. Récupération de la date cible en Config!B12
    const valeurB12 = ongletConfig.getRange("B12").getValue();
    
    if (!(valeurB12 instanceof Date)) {
      // Si la date est invalide ou vide, on masque tout par sécurité
      ongletsAControler.forEach(nomOnglet => {
        let onglet = ss.getSheetByName(nomOnglet);
        if (onglet) onglet.hideSheet();
      });
      ongletConfig.hideSheet();
      return; 
    }

    // Comparaison des dates (sans les heures)
    const aujourdhui = new Date();
    aujourdhui.setHours(0,0,0,0);
    
    const dateAffichage = new Date(valeurB12);
    dateAffichage.setHours(0,0,0,0);

    // 3. Application des visibilités selon le calendrier
    if (aujourdhui >= dateAffichage) {
      // Date atteinte ou dépassée -> Affichage des matières
      ongletsAControler.forEach(nomOnglet => {
        let onglet = ss.getSheetByName(nomOnglet);
        if (onglet) onglet.showSheet();
      });
    } else {
      // Date non atteinte -> Masquage strict des matières
      ongletsAControler.forEach(nomOnglet => {
        let onglet = ss.getSheetByName(nomOnglet);
        if (onglet) onglet.hideSheet();
      });
    }

    // L'onglet Config reste masqué de manière permanente pour l'élève
    ongletConfig.hideSheet();

  } catch(e) {
    Logger.log("Erreur lors du contrôle de la date d'affichage : " + e.message);
  }
}

/**
 * Fonction principale permettant de dupliquer le template, filtrer les onglets,
 * puis injecter les données exclusivement dans l'onglet "Config".
 */
function genererDossiersEleves() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    SpreadsheetApp.getUi().alert("Le script est déjà en cours d'exécution. Veuillez patienter.");
    return;
  }

  try {
    const ui = SpreadsheetApp.getUi();

    // =========================================================================
    // CONFIGURATION DES IDENTIFIANTS DE VOS FICHIERS
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
        let cleNettoyee = cle.toString().replace(/\s+/g, ' ').trim().toLowerCase();
        cellMapping[cleNettoyee] = cellule.toString().trim();
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
    let choixDoublon = null;

    // Boucle sur chaque élève du listing
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
        "affichage": listingData[i][9],
        "matricule": listingData[i][10]
      };

      let nom = donneesEleve["nom"] ? donneesEleve["nom"].toString().trim() : "";
      let prenom = donneesEleve["prenom"] ? donneesEleve["prenom"].toString().trim() : "";
      let optionTexte = donneesEleve["option"] ? donneesEleve["option"].toString().trim() : "Sans Option";
      let optionEleveClé = optionTexte.toLowerCase();
      let sessionTexte = donneesEleve["session"] ? donneesEleve["session"].toString().trim() : "Sans Session";

      let nomNouveauFichier = `${nom} ${prenom} - ${optionTexte}`;

      let dossierPEQBiche = obtenirOuCreerDossier(dossierRacineDest, "PEQ Biche");
      let dossierSession = obtenirOuCreerDossier(dossierPEQBiche, sessionTexte);
      let dossierOption = obtenirOuCreerDossier(dossierSession, optionTexte);

      let fichiersExistants = dossierOption.getFilesByName(nomNouveauFichier);
      let fichierCible = null;
      let existeDeja = false;

      if (fichiersExistants.hasNext()) {
        existeDeja = true;
        fichierCible = fichiersExistants.next();
        
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
            return; 
          }
        }
      }

      let nouveauSs;

      if (existeDeja) {
        if (choixDoublon === "ignorer") {
          compteurDoublonsIgnores++;
          continue; 
        } else if (choixDoublon === "ecraser") {
          nouveauSs = SpreadsheetApp.openById(fichierCible.getId());
          compteurEcrases++;
        }
      } else {
        let copieFichier = templateFile.makeCopy(nomNouveauFichier, dossierOption);
        nouveauSs = SpreadsheetApp.openById(copieFichier.getId());
        
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

      // Écriture sécurisée des données de l'élève
      let ongletConfig = nouveauSs.getSheetByName("Config");
      
      if (ongletConfig) {
        for (let cle in cellMapping) {
          let valeur = donneesEleve[cle];
          
          if (valeur !== undefined && valeur !== null && valeur !== "") {
            let celluleCible = cellMapping[cle];
            try {
              let rangeCible = ongletConfig.getRange(celluleCible);
              
              if (valeur instanceof Date) {
                let dateFormatee = Utilities.formatDate(valeur, nouveauSs.getSpreadsheetTimeZone(), "dd/MM/yyyy");
                rangeCible.setValue(dateFormatee);
              } else {
                rangeCible.setValue(valeur);
              }
              
            } catch(e) {
              Logger.log(`⚠️ Erreur d'écriture pour la clé [${cle}] à la cellule [${celluleCible}] : ${e.message}`);
            }
          }
        }

        // --- SÉCURISATION ET VERROUILLAGE DE L'ONGLET CONFIG ---
        try {
          let protections = ongletConfig.getProtections(SpreadsheetApp.ProtectionType.SHEET);
          let protection = protections.length > 0 ? protections[0] : ongletConfig.protect().setDescription('Sécurité Auto Config');
          
          protection.setWarningOnly(false);
          protection.removeEditors(protection.getEditors());
          
          UTILISATEURS_AUTORISES.forEach(email => {
            try { protection.addEditor(email); } catch(err) {}
          });

          // Option de masquage absolu lié à la protection Google Sheets
          if (typeof protection.setHideSheetOnProtection === 'function') {
             protection.setHideSheetOnProtection(true);
          }
        } catch(eProtection) {
          Logger.log(`⚠️ Impossible de verrouiller l'onglet : ${eProtection.message}`);
        }

        // Commande par défaut pour masquer l'onglet Config visuellement à la création
        ongletConfig.hideSheet();
      }

      SpreadsheetApp.flush();
    }
    
    let messageResultat = `Opération terminée !\n\n📊 Résultat :\n- ${compteurSucces} nouveau(x) dossier(s) créé(s).`;
    if (compteurEcrases > 0) messageResultat += `\n- ${compteurEcrases} dossier(s) mis à jour (données écrasées).`;
    if (compteurDoublonsIgnores > 0) messageResultat += `\n- ${compteurDoublonsIgnores} élève(s) ignoré(s).`;
    ui.alert(messageResultat);

  } catch (erreur) {
    SpreadsheetApp.getUi().alert("Une erreur est survenue : " + erreur.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Fonction utilitaire pour trouver ou créer un sous-dossier Drive.
 */
function obtenirOuCreerDossier(dossierParent, nomDossier) {
  const dossiersTrouves = dossierParent.getFoldersByName(nomDossier);
  if (dossiersTrouves.hasNext()) {
    return dossiersTrouves.next();
  }
  return dossierParent.createFolder(nomDossier);
}
