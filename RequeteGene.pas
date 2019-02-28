{$INCLUDE OptionsCompil.txt}

unit RequeteGene;

interface

uses ADODB,DB,Classes,SysUtils,Messages,ManipXml,General,Windows,Registry,TabAsso,Tab_Var;

const
  cstCleConnexionBD= 'Connexion base de donn�es';  // libell� dans la base de registres
  cstRegLicence = 'Licence'; // libell� dans la base de registres
  cstTailleBaseMin = 100;   // v3.4.8c taille base minimum quand le client est param�tr� � 0
  cstIntervalleAvancement = 1/86400;   // 1s = Intervalle entre deux maj de Avancement (en jours = unit� de TDateTime)
//  cstMsgProgression = WM_USER + 1024;

{$IFDEF MODECHRONO}
const
  fichierMouchard ='d:\ADN\R3Web1\temp\Mouchard.txt';
{$ENDIF}


type

  clsEltChrono = class(TElt)   // Liste des m�thodes appel�es
    NomMethode: string;
    Temps: TDateTime;
    NbAppels: integer;
  end;

  excConnexionBD = class(exception);
  excEnVeille = class(exception);
  excReqIncor = class(exception);
  excParIncor = class(exception);
  excSessionNonTrouvee = class(exception);
  excParamConnexion = class(exception);
  TNomTableCompo = array[TCategorie] of string;    // D�clarations de ces types requise si on veut utiliser
  TNomAutreTable = array[TAutreTable] of string;   // l'affectation directe de tableau

  clsReqGenerique = class(TObject)

    private
      CorpsRequete: string;   // m�morise la requ�te apr�s le "Requ�te="
      TauxAvancement: TPourcentageEntier;    // Pourcentage d'avancement de la requ�te (quand il est g�r�)
      TxAvMin: TPourcentageEntier;
      TxAvMax: TPourcentageEntier;   // taux d'avancement maxi de la phase actuelle (champ interne de la prop. TauxAvancementMaxi);
      DerniereMajAvancement: TDateTime;   // top horloge de la derni�re mise � jour avancement

      procedure MajTauxAvancementMaxi(peValeurTaux: TPourcentageEntier);
      function LitTauxAvancementMaxi: TPourcentageEntier;

    protected
      NomParam,ValParam: TabAttrib;  // tableaux des param�tres
      IdProfil: integer;      // n� de profil
      IdUtilisateur: integer; // n� d'utilisateur
// pass� en public      IdRequete: integer;    // n� de requ�te (pour les requ�tes longues) (v3.5.6)
      ReqAdmin: boolean;   // indique s'il s'agit d'une requ�te d'administration
      ReqCnxTelMobile: boolean;   // indique s'il s'agit d'une requ�te de connexion depuis t�l�phone mobile
      DroitsGeneraux: integer;   // Droits g�n�raux lus dans le profil
      Administrateur: boolean;   // indicateur droits Administrateur
      ModifInhibee: boolean;  // vrai quand le profil a �t� d�grad� en consultation seulement (trop de sessions en maj)
      CreationInhibee: boolean;   // vrai quand seule la cr�ation est interdite (taille limite base atteinte)
      ConsultationSeule: boolean;   // vrai quand la session est d'embl�e en lecture seule (interrogation par t�l. mobile)
      OdtActif: integer;      // n� d'ODT actif (= en cours d'enregistrement)
      OdtExecute: integer;    // n� d'ODT en train d'�tre ex�cut� ou simul�
      RequeteAppelante: clsReqGenerique;  // Requ�te qui a cr�� cet objet (sinon nil)
      DcnxPrevue: TDateTime;    // Heure pr�vue de d�connexion (forc�e par un administrateur)
      MotifDeconnexion: string;        // Motif de la d�connexion pr�vue

      procedure ControleVersionDansBD;   // Contr�le compatibilit� de la version inscrite dans la table Parametre
      procedure ControleVersionClient(const peVersionClient: string);   // Contr�le compatibilit� versions client et serveur

    public
      Query: TADOQuery;       // disponible pour toutes les requ�tes
      Command: TADOCommand;   // disponible pour toutes les requ�tes
      FormatsADN: TFormatSettings;
      PileAppels: TStringList;
{$IFDEF MODECHRONO}
      PileChronos: TList;
      TempsPasse: TableauSouple;
{$ENDIF}
{$IFDEF MODEEXE}
      { A supprimer quand le calcul global des trajets sera pass� dans R3Web }
      ProcRafraichProgression: procedure(peNumSession: integer) of object;    // R�serv� � MaintenanceR3Web.exe
{$ENDIF}
      NumSession: integer;    // n� de session
      NumVue: integer;        // n� de vue courante sur le module client
      IdRequete: integer;    // n� de requ�te (pour les requ�tes longues) (v3.5.6)
      ConnexionBD: TADOConnection;
      NomTableCompo: TNomTableCompo;  // nom des tables utilis�es par la requ�te (seront variables)
      NomAutreTable: TNomAutreTable; // Nom des tables hors composants (pourront changer en mode ODT)
      TailleMaxReponse: integer;   // taille maxi cha�ne
      DureeMaxiRequete: integer;   // dur�e maximum th�orique d'une requ�te
      Environnement: string;      // nom de l'environnement logique
      EnregHistorique: boolean;   // indicateur "enregistrer les op�rations dans l'historique"
      IgnorerAlerte: boolean;     // indicateur "ignorer les avertissements"
      IdSite: string;          // valeur de l'ident du site concern� par la requ�te
      NumAction: integer;    // n� de l'action ex�cut�e (mode direct ou ODT)
      Options: word;          // options d'ex�cution cod�es sur 16  (flags)
      NumActionContexte: integer;    // n� de l'action m�re (contexte) (3.5.3a)

      Password: string;    // champs ajout�s pour se passer de la lecture registre dans le contexte de l'import
      Provider: string;
      UserId: string;
      NomBase: string;
      DataSource: string;
      AuthentifiantWindows: string;    // 3.6.0 SSO : identifiant windows (non vide si IIS configur� en authentification Windows et r�seau local)

      property TauxAvancementMaxi: TPourcentageEntier read LitTauxAvancementMaxi write MajTauxAvancementMaxi;

      constructor Create(peAppelant: clsReqGenerique); virtual;

      destructor Destroy; override;

      procedure EntreeProc(pLibelProc: string);
      { Empile le nom de la proc�dure en cours }
      procedure SortieProc;
      { D�sempile le nom de la proc�dure en cours }

      function ValeurParam(peNomParam: string): string; overload;
      { Donne la valeur d'un param�tre transmis � la requ�te et stock� dans (NomParam,ValParam) }

      function ValeurParam(peNomParam: string; peValeurDefaut: integer): integer; overload;
      { Idem pr�c�dente mais renvoie un entier en le contr�lant }

      procedure LoginBD;
      { Connexion � la base de donn�es }

      procedure ChercheEnvironnement(peReg: TRegistry);

      procedure MajNomTablesPermanentes;
      
      procedure Initialise(peChRequete: string); virtual;
      { Appel�e par TOUTES les requ�tes sauf la connexion initiale, les requ�tes d'authentification et RqDiagnostic}

      function RemplitMessage (peTypeMessage: tTypeMessage; peContenu: string): string; overload;
      { Renvoie une chaine de la forme <message type = "peTypeMessage">peContenu</Message> }
      function RemplitMessage (peTypeMessage: tTypeMessage; peContenu: string;
       peNomAttrib: array of string; peValAttrib: array of string): string; overload;

      function AffichePileAppels: string;
      { Affiche la pile des appels de proc�dures et fonctions }

      function RecupereIdent: integer;
      { Donne l'identifiant automatiquement g�n�r� par la derni�re requ�te de cr�ation }

      procedure SQLRecupereIdent(var peTexteSQL: TStringList; peNomVarSQL: string;
       peDeclarationVar: boolean = true);
      { script SQL des commandes permettant de r�cup�rer le dernier identifiant g�n�r� }

      procedure Finalise; virtual;
      { Remet l'indicateur "requ�te en cours" � 0 }

      procedure AnnuleTransactions;
      { Annule toutes les transactions en cours }

      function RemplitItem (peNom: string; peCategorie: char; peNum: integer = -1;
                            peInfo: integer = 0): string; overload;
      { Renvoie une cha�ne XML contenant les infos de base d'un composant }

      function RemplitItem (peNom: string; peCategorie: char; peNum: integer;
       peInfo: string): string; overload;

       function DoubleQuotes(peChaine: string; peCarDelim: char = '''';
        peEncadrerChaine: boolean = false): string;
      { Double les "'" (ou autre caract�re (v3.4.3) dans une cha�ne pour pouvoir la soumettre
         comme param�tre dans une requ�te SQL ou g�n�rer du CSV}

      function TraiteException(peExc: exception): string;
      { Pr�pare le message � renvoyer au client en fonction de l'exception peExc }

      procedure LigneeItemLieux(peCateg: TCategorie; peIdObj: integer;
       var psRangDansArbre: string; var psLignee: string; var psDroit: TDroit);
      { Renvoie la liste des lieux anc�tres d'un lieu ou �quipement }

      procedure OrdreEtDroitLieu(              // v3.5.2b
        peCateg: TCategorie; peIdObj: integer; var psOrdre: string; var psDroit: TDroit);
      { Recherche le droit d'acc�s et le champ Ordre (si lieu simple) d'un lieu donn� }

      function DansLignee(peCategorie: TCategorie;peIdent: integer;
       peLignee: string): boolean;
      { Renvoie vrai si l'�l�ment (peCategorie,peIdent) est pr�sent dans la lign�e peLignee }

      function CritereSessionPerimee(peTopHorloge: TDateTime; peDelaiVeille: integer;
       peDureeMaxSession: integer; peDureeMaxRequete: integer)
       : string;

      function CritereSessionActive(peTopHorloge: TDateTime; peDelaiVeille: integer;
       peDureeMaxRequete: integer)
       : string;

      function CritereSessionEnVeille(peTopHorloge: TDateTime; peDelaiVeille: integer;
       peDureeMaxSession: integer; peDureeMaxRequete: integer)
       : string;
      { Ces fonctions utilisent les "settings" formatsADN pour les s�parateurs }

      function vpFormatOK
       (const peValeur,peFormat: string;
        var psValeurFormatee: string;
        var psMsg : string): boolean;
      { V�rifie peValeur selon peFormat renvoie TRUE si OK sinon FALSE + psValeurFormatee + psMsg }

      procedure LitParamChaine(peTNumPar: array of integer; peTValeurDefaut: array of string;
       var psValParam: tabAttrib);
      { Lit des param�tres g�n�raux de type cha�ne dont les noms sont dans peTNomsPar }

      procedure LitParamEntiers(peTNumPar: array of integer; peTValeurDefaut: array of integer;
       var psValParam: tabDynEntier);
      { Variante de la pr�c�dente avec des valeurs enti�res }

      function AjouteActionGlobale(peOperation: TOperationR3Web): integer;
      { Ajout d'une action symbolique globale � un site }

      function AjouteActionCablage(
       peCodeOper: TOperationR3Web;   // code op�ration
       peCategorie: TCategorie;     // code cat�gorie d'objet
       peIdObjet: integer;       // identifiant d'objet
       peNomObjet: string;        // nom de l'objet
{       peActionContexte: integer = 0;    // supprim� en v3.5.3a - tient compte de NumActionContexte � la place }
       peComm: string = '';      // Commentaire �ventuel
       peNbCnx: integer = 0;   // Nombre de connexions pour les op�rations autres que maj de composant
       peNomCncDep: string = '';   // nom du premier connecteur de d�part
       peIdCncDep: integer = 0;   // identifiant du premier connecteur de d�part
       peCnxDep: string = '';   // premi�re connexion de d�part ou bien premier fil de fonction retir�
       peNomDerCncDep: string = '';   // nom du dernier connecteur de d�part ou bien dernier fil de fonction retir�e
       peDerCnxDep: string = '';   // premi�re connexion de d�part
       peNomCncArr: string = '';   // nom du premier connecteur d'arriv�e
       peIdCncArr: integer = 0;   // identifiant du premier connecteur d'arriv�e
       peCnxArr: string = '';   // premi�re connexion d'arriv�e
       peLigneeDepart: TStringList = nil;   // LT [+ GE] + eqt de d�part s'il s'agit d'une op�ration �labor�e
       peCategLigneeDepMax: TCategorie = eEquipement;    // Cat�gorie du dernier �l�ment de la lign�e (pour d�c�blage sur tout un GE ou tout un LT)
       peIdEqtDep: integer = 0;   // identifiant de l'�quipement de d�part
       peLigneeArrivee: TStringList = nil;     // LT [+ GE] + eqt d'arriv�e si l'op�ration en poss�de
       peIdEqtArr: integer = 0;   // identifiant de l'�quipement d'arriv�e
       peFilCabFonc: integer = 0;    // premier fil de c�ble c�bl� ou de fonction
       peLongueur: integer = -1;    // longueur (de brassage)
       peValeursProprietes: TIdentValeurFormat = nil)
       : integer;       // renvoie le num�ro d'action g�n�r�e
       { Ajout d'une action �labor�e (c�blage, placement) dans l'historique - utilis�e pour les op�rations de cablage et placement }

      procedure SQLAjouteActionCompo(
       var peTexteSQL: TStringList;  // Commandes SQL � mettre � jour
       peCodeOper: TOperationR3Web;   // code op�ration
       peCategorie: TCategorie;     // code cat�gorie d'objet
       peIdObjet: integer;       // identifiant d'objet
       peNomObjet: string;     // nom de l'objet
       (* peActionContexte: integer = 0;    // supprim� en v3.5.3a - tient compte de NumActionContexte � la place } *)
       peComm: string = '';      // Commentaire �ventuel
       peNomLt: string = '';    // nom du LT d'appartenance si l'objet modifi� est un �quipement
       peNomGe: string = '');   // nom du GE d'appartenance si l'objet modifi� est un eqt dans un GE

      { Ajout d'une action dans l'historique (ne fait que g�n�rer le texte SQL)
        pour les modifications de composant ou autres actions n'ayant ni param�tres d�part ni arriv�e }

      procedure SQLAjouteCreationCompo(
       var peTexteSQL: TStringList;  // Commandes SQL � mettre � jour
       peCategorie: TCategorie;     // code cat�gorie d'objet
       peNomVarSQL: string;       // nom de la variable SQL contenant l'identifiant d'objet
       peNomObjet: string;     // nom de l'objet cr��
      {peActionContexte: integer = 0;    // supprim� en v3.5.3a - tient compte de NumActionContexte � la place }
       peComm: string = '';      // Commentaire �ventuel
       peNomLt: string = '';    // nom du LT d'appartenance si l'objet cr�� est un �quipement
       peNomGe: string = '');    // nom du GE d'appartenance si l'objet cr�� est un eqt dans un GE

      { Analogue � la pr�c�dente, mais en fournissant un nom de variable � la place de peIdObjet }

      function TrouveEntier(peValCherchee: integer; peTabDyn: TabDynEntier;
       var psIndiceTrouve: integer)
       : boolean;
      { Retrouve une valeur dans un tableau dynamique d'entiers  }

      function RecenseEntier(peValCherchee: integer; var pesTabDyn: TabDynEntier)
       : boolean;
      { Cherche une valeur dans un tableau dynamique
      et l'ajoute au tableau si elle n'est pas trouv�e }

      function IdToNom(const peCateg: TCategorie; peIdent: integer; peRendreNonXml: boolean = false): string;
      { Selon la cat�gorie, r�cup�re un nom en fonction d'un identifiant }

      function TrouveTypeComposant(const peCateg: TCategorie; const peIdent: integer): integer;

      function NomComposant: string;
      { Nom d'un composant en fonction de son n� et de sa cat�gorie (NPC avec NomcomposantUnique) }

      function NumeroSite(peNomSite: string): integer;
      { Num�ro d'un lieu de niveau 1 en fonction de son nom }

      { Renvoie l'identifiant d'un c�ble, LT, terminaison, BN ou fonction }
      function NumeroComposant(const peCateg: TCategorie; const peNom: string;
       peNumSite: integer; var psNumType: integer; var psNumLieu: integer)
       : integer;

      function NumeroEquipement(peId_cnc: integer): integer;
      { Renvoie le num�ro d'�quipement auquel le connecteur peId_cnc appartient, 0 si pas trouv� }

      function IdTypeToNom(peId_type: integer): string;
      { Donne un nom de type en fonction de son identifiant }

      function LigneeCnc(peId_cnc: integer; var psLigneeAffichee: string; var psTypeCnc: string): string;
      { Nom connecteur et liste LT/GE/Eqt + type � partir de l'identifiant d'un connecteur }

      function LigneeCncXml(peId_cnc: integer; var psTailleGroupe: integer; var psNbGroupes: integer): string;
      { Cha�ne XML contenant les noms LT,GE,Eqt,Cnc � partir de l'identifiant d'un connecteur }

      procedure TrouveNomLtGe(peId_lieu: integer; var psNomLt: string; var psNomGe: string);
     { Recherche du nom d'un LT ou du nom d'un GE + son LT d'appartenance }

      function DroitAccesLieuSimple(peNumLieu: integer; peNiveauLieu: integer): TDroit;
      { Droit d'acc�s au lieu en fonction de son num�ro }

      function DroitAccesLieuTechnique(peNumLieu: integer; peCateg: TCategorie): TDroit;
      { Droit d'acc�s � un local technique ou un groupe d'�quipements en fonction de son num�ro }

      { Recherche des droits et du num�ro d'un local technique ou d'un GE             }
      { La proc�dure sert aussi de test d'existence : renvoie psNumero = 0 si pas trouv� }
      function NumeroLieuTechnique(peNumSite: integer; peNomSite: string;
       peNomLt: string; peNomGe: string;
       var psDroit: TDroit): integer;

      function DroitAccesEquipement(peNumEqt: integer): TDroit;
      { Droit d'acc�s � un �quipement }

      procedure CreerTablesTempo;
      { Cr�ation de tables temporaires en copiant les tables originales }

      procedure SupprimerTablesTempo;
      { Suppression des tables temporaires li�es � une session et �ventuellement une vue }

(*    Suppr. v3.4.5a  function NomStandardCnx(peCnx: integer; peTailleGroupe: integer;
       peCestLeFinFin: boolean = false): string;
      { Renvoie une d�signation standard de connexion ou de fil de c�ble }
*)

      function NomCnx(peCnx: integer; peNbCnx: integer; peTailleGroupe: integer;
       peNbGroupes: integer; var psGroupagePossible: boolean; peRetourVidePossible: boolean = true)
       : string;

      function LibelleSerieCnxOuFils(pePremierFil: integer; peNbCnx: integer;
       peTailleGroupe: integer; peNbGroupes: integer)
       : string;
       { D�signation standard d'une s�rie de connexions ou de fils de c�ble pour affichage }

      function LitRegistreADN(peNomCle: string; peNomValeur: string; var psValeurLue: string)
       : boolean;
      { Lecture d'une valeur de registre dans HKEY_LOCAL_MACHINE\SOFTWARE\ADN\R3Web }

      function DonneCheminAcces(peNomValeurRegistre: string; peCheminComplet: boolean = true): string;     // v3.5.6 refondue 3.6.5

      function Diagnostic(peChRequete: string): string;
      { Proc�dure appel�e � partir du module Flash sp�cial de diagnostic }

      function DateR3WebClient(peDate: TDateTime)   // date format Delphi
      : string;            // Renvoie une cha�ne JJ/MM/AAAA

      function DateR3WebServeur(peChaineDate: string)   // cha�ne en format JJ/MM/AAAA
      : integer;     // Renvoie une valeur stockable en base

      function DonneNomNiveauSite: string; overload;

      function DonneNomNiveauSite(var psGenreGr: TGenreGr): string; overload;

      procedure ChargeXmlToTab(pesoTa: clsTabAsso; peBalise, peAtt, peXml: string);

      procedure AncetresConnecteur(peIdCnc: integer;
       var psNomLt: string; var psNomGe: string; var psNomEqt: string);
      // Donne les noms de LT, de GE et d'�quipement ou terminaison d'un connecteur

      procedure LitCleProtec(
       peAppelParImport: boolean;   // true si appel par Import
       var psCheminAccesServeur: string;  // chemin d'acc�s au serveur
       var psNbAccesMaj: integer;   // nombre maxi d'acc�s simultan�s en mise � jour
       var psNbAccesCon: integer;   // nombre maxi d'acc�s simultan�s en consultation
       var psNbMilliers: integer;   // taille base de donn�es maxi
       var psCodeClient: integer);  // n� licence client
      // D�code les caract�ristiques de la licence client

      function ControleTailleBase(
       peTailleBase: integer;   // taille autoris�e
       peCodeClient: integer;   // code client (pour shunter certains contr�les)
       var psOccupation: integer;  // nb de (pseudo-)connecteurs ou de fonctions selon ce qui approche le plus de peTailleBase
       var psCategCause: TCategorie)  // indicateur cause de d�passement : T si terminaisons, F si fonctions
       : boolean;   // true si la taille autoris�e est d�pass�e

      procedure ControleNbAcces(
       var pesProfilModif: boolean;  // vrai si profil de la session est en modification
       peNbAccesMaj,peNbAccesCon: integer;   // nombres d'acc�s autoris�s en mise � jour et en consultation
       peTopHorloge: TDateTime;    // top horloge actuel
       peDelaiVeille,peDureeMaxRequete: integer);   // valeurs des param�tres g�n�raux correspondants
      // contr�le le nombre d'acc�s en mise � jour et en consultation

      function LitCheminImportExport(peNumeroParametre: integer): string;

      function ListeLieuxNiveau1: string;   // donne la liste des lieux de niveau 1 (sites)

      function CreeADOQuery(peDelaiInfini: boolean = true): TADOQuery;   // cr�ation d'une requ�te

      function NomComposantUnique: string;   // donne un nom unique � attribuer � un composant

      procedure SupprimeActionPrevue(           // Suppression d'action d'ODT
       peActionPrincipale: TOperationR3Web;     // code de l'action principale
       peCategorie: TCategorie = eLieuSimple;  // cat�gorie d'objet si peActionPrincipale = eCreerComposant ou eModifierComposant ou eSupprimerComposant
       peContexteSupprManu: boolean = false);  // true si l'appel vient de clsOperation.SupprActionOdt (v3.5.3a)

      function DroitModifOdt(   // Donne le  droit de modification sur un ODT
       peProfilOdt: integer)  //  Profil de l'ODT
       : boolean;   // Renvoie true si l'ODT est modifiable, false sinon

      procedure MajAvancement(
       pePourcentage: TPourcentageEntier;
       peLibelle: string = '';
       peMajTxAvMin: boolean = true);    // false s'il ne faut pas maj TauxAvDebutPhase

      procedure AjouteAvancement(peProportionAvancementPartiel: real);

{$IFDEF MODECHRONO}
      procedure Mouchard(peTexte: string);
{$ENDIF}
    end;

  { v3.6.0 - gestion des param�tres longs }
  type clsEnvoiLongParam = class(clsReqGenerique)
  end;

const
  cstPrefixeLieuSimple = #15;
  cstPrefixeLocalTechnique = #16;
  cstPrefixeGroupe = #17;
  cstPrefixeEquipement = #18;
  cstPrefixeTerminaison = #19;
  cstPrefixeComposant = cstPrefixeLieuSimple;

//  cstSeparateur = '|';  // Pour les messages d'erreur
  lbErrYourLicenseIsNotValidForEnglishVersion = 'Your license is not valid for the English version';
  lbErrVotreLicenceNestPasValidePourLaVersionFrancaise = 'Votre licence n''est pas valide pour la version fran�aise';

implementation

uses ResStr, Balises, Session, StrUtils, ExportXls;

{$IFDEF MODECHRONO}
{ ---------------------------------------------------------------------------- }
function CompareEltChrono(var peChrono1,peChrono2: clsEltChrono): tResuCompar;
{ Fonction de comparaison pour tableauSouple d'�l�ments de type clsTexteCnc }
{ Renvoie eInferieur si peCnc1 < peCnc2 ,
           eSuperieur si peCnc1 > PeCnc2
           eEgal sinon }
{ ---------------------------------------------------------------------------- }

begin
  if peChrono1.NomMethode < peChrono2.NomMethode then
    result:= eInferieur
  else
    if peChrono1.NomMethode > peChrono2.NomMethode then
      result:= eSuperieur
    else
      result:= eEgal;
end;
{$ENDIF}

{ --------------------------------------------------------------------------------------- }
constructor clsReqGenerique.Create(peAppelant: clsReqGenerique);
{ --------------------------------------------------------------------------------------- }
{ Rappel : Si une exception est d�clench�e lors de l'ex�cution d'un constructeur appel�
dans une r�f�rence de classe, le destructeur Destroy est appel� automatiquement
pour d�truire l'objet inachev�. }
begin
  if peAppelant = nil then
  begin
    PileAppels:= TStringList.Create;
{$IFDEF MODECHRONO}
    PileChronos:= TList.Create;
    TempsPasse:= TableauSouple.Create(clsEltChrono);
    TempsPasse.FoncCompTs:= @CompareEltChrono;   // fonction de comparaison qui doit optimiser les recherches
{$ENDIF}
    ConnexionBD:= TADOConnection.Create(nil);
    ConnexionBD.LoginPrompt:= true;
    ConnexionBD.CommandTimeout:= 0;      // v3.5.0 correction bug � l'entr�e dans R3Web "Connexion � la base de donn�es impossible - d�lai d�pass�"
    Command:= TADOCommand.Create(nil);   // disponible pour toutes les requ�tes
    Command.Connection:= ConnexionBD;
    Command.CommandTimeout:= 0;        // v3.4.8 (861)
    // Si on met := DureeMaxiRequete �a fait toujours des erreurs "D�lai d�pass�"
    Command.ParamCheck:= false;     // v3.5.0 (1015)
    Query:= CreeADOQuery();       // disponible pour toutes les requ�tes (modif 3.5.0)
    RequeteAppelante:= nil;    // requ�te directement appel�e depuis une action Web
    EnregHistorique:= true;   // enregistrer les actions dans l'historique
    IgnorerAlerte:= false;    // montrer les messages d'avertissement en annulant la transaction
    IdSite:= '';
    OdtExecute:= 0;
    NumActionContexte:= 0;   // v3.5.3a (1133) - non modifi� dans le cas o� peAppelant <> nil
    IdUtilisateur:= 0;      // 3.6.0 ajout� pour le SSO

    Password:= '';    // champs ajout�s pour se passer de la lecture registre dans le contexte de l'import
    Provider:= '';
    UserId:= '';
    NomBase:= '';
    DataSource:= '';
{$IFDEF MODEEXE}
    ProcRafraichProgression:= nil;    // v3.6.0 - valeur par d�faut (utilis� pour MaintenanceR3Web)
{$ENDIF}

{$IFDEF MODECHRONO}
    if fileExists(fichierMouchard) then
      DeleteFile(fichierMouchard);
{$ENDIF}
  end
  else   // C'est un objet d�riv� de clsReqGene qui cr�e un autre objet d�riv� de clsReqGene
  begin
  { R�cup�ration des infos issues de l'objet appelant }
    PileAppels:= peAppelant.PileAppels;
{$IFDEF MODECHRONO}
    PileChronos:= peAppelant.PileChronos;
    TempsPasse:= peAppelant.TempsPasse;
{$ENDIF}
    NomParam:= peAppelant.NomParam;  // Attention : ce n'est pas une copie mais LE MEME tableau dynamique
    ValParam:= peAppelant.ValParam;  // idem
    CorpsRequete:= peAppelant.CorpsRequete;   // m�morise la requ�te apr�s le "Requ�te="
    NumSession:= peAppelant.NumSession;    // n� de session
    NumVue:= peAppelant.NumVue;
    IdProfil:= peAppelant.IdProfil;      // n� de profil
    IdUtilisateur:= peAppelant.IdUtilisateur;  // n� d'utilisateur
    IdRequete:= peAppelant.IdRequete;     // v3.5.6
    Administrateur:= peAppelant.Administrateur;
    DroitsGeneraux:= peAppelant.DroitsGeneraux;
    ModifInhibee:= peAppelant.ModifInhibee;
    CreationInhibee:= peAppelant.CreationInhibee;   // v3.6.0 (1242)
    OdtActif:= peAppelant.OdtActif;      // n� d'ODT actif
    OdtExecute:= peAppelant.OdtExecute;
    TailleMaxReponse:= peAppelant.TailleMaxReponse;
    DureeMaxiRequete:= peAppelant.DureeMaxiRequete;
    RequeteAppelante:= peAppelant;    // objet requ�te qui a cr�� cet objet
    Environnement:= peAppelant.Environnement;
    { Les composant standard Command ne sont pas ceux de l'appelant : ils sont cr��s pour �tre d�di�s � l'objet }
    { On r�cup�re la connexion � la BD, initialis�e par l'appelant }
    ConnexionBD:= peAppelant.ConnexionBD;
    Command:= TADOCommand.Create(nil);   // disponible pour toutes les requ�tes
    Query:= CreeADOQuery();       // disponible pour toutes les requ�tes (v3.5.0c passage � CreeADOQuery (1015))
    Command.Connection:= ConnexionBD;
    Command.CommandTimeout:= 0;     // Si on met := DureeMaxiRequete �a fait toujours des erreurs "D�lai d�pass�"
    NomTableCompo:= peAppelant.NomTableCompo;     // nom des tables utilis�es par la requ�te (seront variables)
    NomAutreTable:= peAppelant.NomAutreTable;     // tables variables autres que celles des objets
    EnregHistorique:= peAppelant.EnregHistorique;
    IgnorerAlerte:= peAppelant.IgnorerAlerte;
    IdSite:= peAppelant.IdSite;
    DcnxPrevue:= peAppelant.DcnxPrevue;
    TailleMaxReponse:= peAppelant.TailleMaxReponse;
    DureeMaxiRequete:= peAppelant.DureeMaxiRequete;
    DerniereMajAvancement:= peAppelant.DerniereMajAvancement;
    TauxAvancement:= peAppelant.TauxAvancement;
    TauxAvancementMaxi:= peAppelant.TauxAvancementMaxi;

    Password:= peAppelant.Password;    // champs ajout�s pour se passer de la lecture registre dans le contexte de l'import
    Provider:= peAppelant.Provider;
    UserId:= peAppelant.UserId;
    NomBase:= peAppelant.NomBase;
    DataSource:= peAppelant.DataSource;
{$IFDEF MODEEXE}
    ProcRafraichProgression:= peAppelant.ProcRafraichProgression;    // 3.6.0
{$ENDIF}

  end;
  NumAction:= 0;    // l'action li�e � cette requ�te n'existe pas (encore)
  Options:= 0;      // options d'ex�cution: par d�faut aucune

{ liste des valeurs TFormatSettings sur un poste en locale fran�aise
  CurrencyFormat: 3
  NegCurrFormat: 8
  ThousandSeparator: �
  DecimalSeparator: ,
  CurrencyDecimals: 2
  DateSeparator: /
  TimeSeparator: :
  ListSeparator: ;
  CurrencyString: �
  ShortDateFormat: dd/MM/yyyy
  LongDateFormat: dddd d MMMM yyyy
  TimeAMString:
  TimePMString:
  ShortTimeFormat: hh:mm
  LongTimeFormat: hh:mm:ss
  TwoDigitYearCenturyWindow: 0
}
  with FormatsADN do
  begin
    DecimalSeparator:= '.';
    DateSeparator:= '/';
    ShortDateFormat:= 'dd/mm/yyyy';
    timeSeparator:= ':';
    LongTimeFormat:='hh:mm:ss';
    ShortTimeFormat:= 'hh:mm';
  end;
end;


destructor clsReqGenerique.Destroy;
begin
  if requeteAppelante = nil then   // Si la requ�te n'a pas �t� cr��e depuis une autre requ�te
  begin
//    ConnexionBD.Close;   // inutile ?
    ConnexionBD.Free;    // utile en particulier pour l'import
  end;
  Command.Free;    // Command et Query sont propres � cette instance d'objet, donc on doit les lib�rer
  Query.Free;
(*  setLength(NomParam,0);    // devrait lib�rer la m�moire utilis�e ?
  setLength(ValParam,0);
  Normalement pas n�cessaire *)
  if RequeteAppelante = nil then
  begin
    PileAppels.Free;
  // sinon, ne pas d�truire la pile qui en fait celle de l'appelant !
{$IFDEF MODECHRONO}
    PileChronos.Free;
    TempsPasse.Free;
{$ENDIF}
  end;
  inherited Destroy;
end;


{ ---------------------------------------------------------------------------------------------- }
procedure clsReqGenerique.ChercheEnvironnement(peReg: TRegistry);
{ D�tection automatique de l'environnement s'il n'est pas fourni }
{ Contexte: peReg est d�j� cr�� et initialis� - Environnement est � vide }
{ Sortie: s'il y a une cl� unique correspondant � l'environnement dans la base de registres,
    Environnement est mis � jour avec le nom de cette cl�
    Sinon Exception externe }
{ ---------------------------------------------------------------------------------------------- }

var
  lCles: TStringList;
  indS: integer;
  environnementTrouve: boolean;
  sousCle: TRegistry;

begin
  lCles:= TStringList.Create;
  sousCle:= TRegistry.Create;
  try
    if not peReg.OpenKey(cstRegR3Web,false) then
      EXCEPTIONINTERNE(defErr300);
    peReg.GetKeyNames(lCles);
    sousCle.RootKey:= HKEY_LOCAL_MACHINE;
    sousCle.access := KEY_READ;
    { Compte le nombre de cl�s qui sont structur�es comme des environnements
     (cela �vite la cl� DEBUG et d'autres �ventuelles futures cl�s) }
    environnementTrouve:= false;
    for indS:= 0 to pred(lCles.Count) do
    begin
	  //v3.5.0a MC lbCleConnexionBD --> cstCleConnexionBD, lbRegCheminsDAcces -->cstRegCheminsDAcces
      if sousCle.OpenKey(cstRegR3Web+lCles.Strings[indS]+'\'+cstCleConnexionBD,false) then
      begin
        sousCle.CloseKey;
        if sousCle.OpenKey(cstRegR3Web+lCles.Strings[indS]+'\'+cstRegCheminsDAcces,false) then
        begin
          sousCle.CloseKey;
          if environnementTrouve then   // Si 2 environnements ont �t� d�finis
            EXCEPTIONEXTERNE(lbErrEnvLogManquant);  // Il faut en fournir un en param�tre
          environnementTrouve:= true;
          Environnement:= lCles.Strings[indS];
        end;
      end;
    end;

    if not environnementTrouve then
      EXCEPTIONINTERNE(defErr303);
  finally
    peReg.CloseKey;   // Indispensable sinon pb � la lecture suivante
    sousCle.Free;
    lCles.Free
  end;
end;

{ ---------------------------------------------------------------------------------------------- }
procedure clsReqGenerique.LoginBD;
{ Connexion � la base de donn�es }
{ ---------------------------------------------------------------------------------------------- }

var
//  dataSource,nomBase,nomUtil,serveur,mdp: string;
// On utilise directement les champ publics
// (en import ils peuvent �tre d�j� � jour suite au lancement de la requ�te RqLitRegistreR3Web � la DLL)
  reg: TRegistry;
  paramPassword,paramProvider,paramUserId,paramNomBase,paramDataSource,cheminR3Web: string;


begin
   { Lecture des param�tres de connexion dans la base de registres}
  if not ConnexionBD.Connected then
  begin
    if Provider = '' then
    begin          // Les param�tres de connexion ne sont pas � jour
      reg:= TRegistry.Create;
      try
        reg.RootKey:= HKEY_LOCAL_MACHINE;
        reg.access := KEY_READ;
        if Environnement = '' then
          ChercheEnvironnement(reg);    // Met � jour Environnement si un seul environnement a �t� d�fini
        if reg.OpenKey(cstRegR3Web+Environnement+'\'+cstCleConnexionBD,false) then
          with reg do
          begin
            paramPassword:= ReadString('Password');
            paramProvider:= ReadString('Provider');
            paramUserId:= ReadString('User ID');
            paramNomBase:= ReadString('Initial Catalog');
            paramDataSource:= ReadString('Data Source');
            reg.CloseKey;
          end
        else
          RAISE excParamConnexion.Create(lbErrParamConnexion);

        if reg.OpenKey(cstRegR3Web+Environnement+'\'+cstRegCheminsDAcces,false) then
        begin
          cheminR3Web:= reg.ReadString(cstRegR3Serveur);
          if fileExists(includeTrailingPathDelimiter(cheminR3Web)+'TEMP\'+cstNomFichierFermeture) then
            EXCEPTIONEXTERNE(lbErrAccesR3WebFerme);
        end;

      finally
        reg.Free;
      end;
    end
    else
    begin
      paramProvider:= Provider;
      paramUserId:= UserId;
      paramPassword:= Password;
      paramNomBase:= NomBase;
      paramDataSource:= DataSource;
    end;
    Try
       ConnexionBD.ConnectionString := format(
        'Provider=%s;Persist Security Info=False;User ID=%s;Password=%s;'
        +'Initial Catalog=%s;Data Source=%s',
        [paramProvider,paramUserId,DecodeMdpBase(paramPassword),paramNomBase,paramDataSource]);
       ConnexionBD.LoginPrompt:= false;
       ConnexionBD.CommandTimeout:= 0;     // essai pour r�gler les �checs de connexion (arrive sur ordi MG)
       ConnexionBD.Open;
    Except       // EOleException
       on e: Exception do
         RAISE excConnexionBD.Create(lbErrConnexionBD+sautDeLigne+e.Message)
    end;
  end

end;

{ --------------------------------------------------------------------------------------- }
{ Mise � jour des tables NomTableCompo et NomAutreTable avec les noms des tables permanentes }
{ Mis dans une proc. en 3.6.3 }
procedure clsReqGenerique.MajNomTablesPermanentes;
{ --------------------------------------------------------------------------------------- }
var
  iC: TCategorie;

begin
  EntreeProc(ClassName+'.MajNomTablesPermanentes');

  { v3.6.7 plus simple et plus carr� }
  for iC:= low(TCategorie) to high(TCategorie) do
    NomTableCompo[iC]:= InfoTable[InfoCateg[iC].Table].NomTableReelle;
(*
  { D�termine le nom des tables � utiliser }
  NomTableCompo[eLieuSimple]:= 'Lieu';
  NomTableCompo[eLocalTechnique]:= 'Lieu';
  NomTableCompo[eGroupe]:= 'Lieu';
  NomTableCompo[eEquipement]:= 'Equipement';
  NomTableCompo[eTerminaison]:= 'Equipement';
  NomTableCompo[eConnecteur]:= 'Connecteur';
  NomTableCompo[eFonction]:= 'Fonction';
  NomTableCompo[eCable]:= 'Cable';
  NomTableCompo[eBoiteNoire]:= 'Cable';
  NomTableCompo[eOrdreTravaux]:= 'ODT';
  NomTableCompo[ePlanLieu]:= 'PlanLieu';
  NomTableCompo[eBrassage]:= 'Lien';   // v3.5.4: utilis� pour propri�t�s de brassage
*)

  NomAutreTable[eLien]:= 'Lien';
  NomAutreTable[eParcours]:= 'Parcours';
  NomAutreTable[eDepart]:= 'Depart';
  NomAutreTable[eCablage]:= 'Cablage';
  NomAutreTable[eSymbole]:= 'Symbole';
  NomAutreTable[eValeurProp]:= 'ValeurProp';
  NomAutreTable[eTrajet]:= 'Trajet';
  NomAutreTable[eExtTrajet]:= 'ExtTrajet';
  NomAutreTable[eTrajet_Famille]:= 'Trajet_Famille';

  SortieProc
end;


{ --------------------------------------------------------------------------------------- }
procedure clsReqGenerique.Initialise(
 peChRequete: string);     // param�tres de la requ�te
{ --------------------------------------------------------------------------------------- }

var
  chValParam,chMdp,debutNom,inutil1,versionClient,chaineCnxActiviteR3Web,chCnxBd: string;
  tabNomInfoSession,tabValInfoSession,tabNom,tabVal: TabAttrib;
  presente,profilModif,sessionSSO,traiterRequete,tailleBaseDepassee: boolean;
  numSessRecu,compteParam,nbMilliers,codeClient,nbAccesMaj,nbAccesCon,totaltrouve: integer;
  iL,ptrPremierParam,positFinBaliseParam,aFournir: integer;
  topHorloge,topHorlVeille,topHorlSession,valTopRequete: TDateTime;
  ParamGenNum: TabDynEntier;
  indCat: TCategorie;
  indAutre: TAutreTable;
  QSession: TADOQuery;
  codeAccesSession: char;
  categCause: TCategorie;
  lstLongParam: TStringList;
  nomFicLongParam,cheminR3Web,contenuSession: string;

begin
  EntreeProc(className+'.Initialise');
  { M�morise la requ�te pour pouvoir l'afficher en cas d'erreur }
  CorpsRequete:= peChRequete;
  contenuSession:= ValChampXml(defBalInfoSession,peChRequete,ptrPremierParam,tabNomInfoSession,tabValInfoSession,presente);
  // contenuSession ne peut contenir en fait que la balise <mdp>, les autres infos �tant dans les attributs
  // 3.6.0: pour m�moriser la partie <session> de la requ�te afin de la reconstruire
  if presente then
  begin
    Environnement:= ValeurAttribut(defAttEnvLog,tabNomInfoSession,tabValInfoSession);
    ReqAdmin:= upperCase(ValeurAttribut(defAttAdmin,tabNomInfoSession,tabValInfoSession)) = 'O' ;
    NumVue:= StrToIntDef(ValeurAttribut(defAttVue,tabNomInfoSession,tabValInfoSession),0);   // d�plac� ici (1022)
    IdRequete:= StrToIntDef(ValeurAttribut(defAttIdReq,tabNomInfoSession,tabValInfoSession),0);   // 3.5.6 - pr�sent si requ�te longue ou envoi de param�tres longs

    versionClient:= ValeurAttribut(defAttVersion,tabNomInfoSession,tabValInfoSession);

    if versionClient <> '' then   // on ne contr�le que si versionClient est fourni (donc pas dans contexte import,majFonctionsR3Web, etc.)
      ControleVersionClient(versionClient);

    numSessRecu:= StrToIntDef(ValeurAttribut(defAttNumero,tabNomInfoSession,tabValInfoSession),0);
    if numSessRecu = 0 then
      RAISE excReqIncor.Create(lbErrReqIncor+'('+defAttNumero+')');  // pas d'attribut n� de session

    chCnxBd:= ValChampXml(defBalConnexionBD,peChRequete,tabNom,tabVal,presente);
    if presente then
    begin     // Ceci va dispenser LoginBD de consulter la base de registres pour avoir les param�tres de connexion
      Provider:= RendNonXml(ValeurAttribut(defAttProvider, tabNom, tabVal));
      UserId:= RendNonXml(ValeurAttribut(defAttUserId, tabNom, tabVal));
      Password:= HexaVersChaine(ValeurAttribut(defAttPassword, tabNom, tabVal));
      DataSource:= RendNonXml(ValeurAttribut(defAttDataSource, tabNom, tabVal));
      NomBase:= RendNonXml(ValeurAttribut(defAttNomBase, tabNom, tabVal));
    end;

  end
  else
    RAISE excReqIncor.Create(peChRequete);  // pas de champ infosession

  LoginBD;

  { R�cup�re le pointeur sur le composant g�n�ral ADOConnection }
//  Query.Connection:= ConnexionBD;
//  Command.Connection:= ConnexionBD;
  { Valeur d'isolation de transaction par d�faut }
  ConnexionBD.IsolationLevel:= ilReadCommitted;
//  chercher la partie infosession


  { Recherche des param�tres g�n�raux }
  LitParamEntiers([defParDelaiVeille,defParDureeMaxiSession,defParDureeExecRequete,defParTailleMaxReponse],
   [1200,43200,180,15000],ParamGenNum);
  DureeMaxiRequete:= ParamGenNum[2];
  TailleMaxReponse:= ParamGenNum[3];

  NumSession:= numSessRecu;
  chMdp:= ValChampXml(defBalMdp,contenuSession);
  QSession:= CreeADOQuery();
  try
//    QSession.Connection:= ConnexionBD;

    with QSession do
    begin
      SQL.Clear;
      SQL.Add('Select S.Id_uti,S.UtilUnique,S.Id_Odt,S.TopRequete,S.DcnxPrevue,S.SSO,');
      SQL.Add('S.Preferences,S.Contexte,U.Id_prof,U.MotPasse,U.Authentifiant,P.Droits,P.Nom from Session S');
      SQL.Add('join Utilisateur U on U.Id_uti = S.Id_uti');
      SQL.Add('join Profil P on P.Id_prof = U.Id_prof');
      SQL.Add(format('where id_ses=%d',[numSessRecu]));
      Open;
      if Eof then
        RAISE excArretTotal.Create(lbErrSessionSupprimee);   // 3.5.3c (1195)
//        EXCEPTIONEXTERNE(lbErrSessionSupprimee);

      { 3.6.0 - Si session SSO, v�rification que l'utilisateur Windows associ� � l'utilisateur R3Web
       est bien celui qui a �t� d�tect� par le serveur Web }
      sessionSSO:= FieldByName('SSO').AsBoolean;
      if sessionSSO and (uppercase(FieldByName('Authentifiant').AsString) <> uppercase(AuthentifiantWindows)) then
        EXCEPTIONINTERNE(defErr223);

      topHorloge:= Date + getTime;
      { pour �viter de multiplier les appels � FieldByName }
      valTopRequete:= fieldByName('TopRequete').asFloat;
      { Test session p�rim�e }
      topHorlVeille:= topHorloge - ParamGenNum[0]/86400;  // temps avant lequel une session est en veille
      topHorlSession:= topHorloge - ParamGenNum[1]/86400;  // temps avant lequel une session est p�rim�e

      if (valTopRequete < topHorlSession)
       or fieldByName('UtilUnique').asBoolean and (valTopRequete < topHorlVeille)
         // La session a fait sa derni�re requ�te avant topHorlSession
         // ou topHorlVeille si elle est en mode utilisateur unique
       or (not fieldByName('DcnxPrevue').IsNull and (fieldByName('DcnxPrevue').asFloat < topHorloge)) then // La session a d�pass� son d�lai de d�connexion
        EXCEPTIONEXTERNE(lbErrSessionPerimee);
      codeAccesSession:= FieldByName('Preferences').AsString[1];

      if self is clsSession then   // si c'est la requ�te RqFinR3 (c'est la seule de classe clsSession qui appelle Initialise)
      begin
        { Mise � jour statistiques d'activit� (852)}
        if LitRegistreADN(cstRegStatActivite,'',chaineCnxActiviteR3Web) then
          with TADOCommand.Create(nil) do
            try
              ConnectionString:= chaineCnxActiviteR3Web;
              try
                CommandText:= format('update Activite set TopFin = %5.6f where NumSession = %d',
                 [topHorloge,NumSession],formatsADN);
                Execute;
              except
                // masque l'erreur d'ex�cution
              end;
            finally
              Free
            end;
      end
      else
      begin
        if chMdp = '' then
        begin
          { Test session en veille }
          if not sessionSSO and (codeAccesSession <> cstProfilTelMobile) and (valTopRequete < topHorlVeille) then
            // Pas de mise en veille pour le SSO ni pour (v3.6.1) la consultation t�l�phone mobile
            RAISE(excEnVeille.Create(lbErrSessionEnVeille));
        end
        else
        begin   // cas o� la session �tait en veille et l'utilisateur a saisi le mot de passe
(* 3.6.6a (1326) : test nuisible          if not sessionSSO and (   *)
          if FieldByName('motPasse').AsInteger <> integer(StrToInt64(chMdp)) then
            // integer(StrToInt64()) traduit les valeurs sup�rieures � High(integer) en valeurs n�gatives
            EXCEPTIONEXTERNE(lbErrMdpIncorrect);
          { R�veil de session: contr�ler la taille de base de donn�es et le nombre d'acc�s simultan�s }
          LitCleProtec(false,inutil1,nbAccesMaj,nbAccesCon,nbMilliers,codeClient);

(* avant la 3.6.5 (1291)      if ControleTailleBase(nbMilliers*1000,codeClient,totaltrouve,categCause) then
           { Contrairement au d�but de session on bloque compl�tement (on pourrait aussi passer en mode cr�ation inhib�e }
            if categCause = eCable then
              EXCEPTIONEXTERNE(lbErrTailleMaxiAtteinte)   // On n'indique rien quand c'est � cause des c�bles ...
            else
              EXCEPTIONEXTERNE(lbErrTailleMaxiAtteinte+format(' (%s)',[InfoCateg[categCause].Code]));
*)
          tailleBaseDepassee:= ControleTailleBase(nbMilliers*1000,codeClient,totaltrouve,categCause);

          profilModif:= codeAccesSession in [cstProfilModif,cstProfilCreationInhibee,cstProfilModifInhibee];
          if not ReqAdmin then   // v3.5.3b (1170)
            ControleNbAcces(profilModif, nbAccesMaj,nbAccesCon,
             topHorloge,ParamGenNum[0],ParamGenNum[2]);   // NB: ProfilModif passera � false s'il n'y a pas assez d'acc�s
          if profilModif then
            if tailleBaseDepassee then // 3.6.5a (1291)
              codeAccesSession:= cstProfilCreationInhibee
            else
              codeAccesSession:= cstProfilModif   // il pouvait �tre � cstProfilModifInhibee
          else  // trop d'acc�s en MAJ
            if codeAccesSession in [cstProfilModif,cstProfilCreationInhibee] then
              codeAccesSession:= cstProfilModifInhibee;
            // Le droit de modification est inhib� cause trop de sessions en maj
        end;
                                                                                           
        { Met � jour les stats d'activit� }
        if LitRegistreADN(cstRegStatActivite,'',chaineCnxActiviteR3Web) then
          with TADOCommand.Create(nil) do
            try
              try
                ConnectionString:= chaineCnxActiviteR3Web;
                CommandText:= format('update Activite set NbRequetes = NbRequetes + 1, TopRequete = %5.6f where NumSession = %d',
                 [topHorloge,NumSession],FormatsADN);
                Execute;
              except
                // masque l'erreur d'ex�cution
              end;
            finally
              Free
            end;
      end;

      ModifInhibee:= codeAccesSession = cstProfilModifInhibee;  // indicateur interdiction de modif pour les requ�tes utilisateur
      CreationInhibee:= codeAccesSession = cstProfilCreationInhibee;
      ConsultationSeule:= codeAccesSession in [cstProfilConsult,cstProfilTelMobile];   // indicateur cr�� en v3.6.1
      DcnxPrevue:= FieldByName('DcnxPrevue').AsFloat;
      if DcnxPrevue <> 0 then
        MotifDeconnexion:= FieldByName('Contexte').AsString;   // on a mis le motif (�ventuel) l�-dedans
      IdProfil:= FieldByName('Id_Prof').asInteger;   // m�morise le profil utilisateur
      DroitsGeneraux:= FieldByName('Droits').AsInteger;
      Administrateur:= (not DroitsGeneraux) = 0;
      // un profil administrateur a $FFFFFFFF ( = -1) comme valeur de Droits
      (* Supprim� v3.4.1a: on peut lancer le module Administration si on a le droit "administrer les plans"
       ou le droit "d�finir les types et les propri�t�s"
       C'est le commencement de session qui contr�le les droits
      if ReqAdmin and not Administrateur then
        // Cela ne peut arriver que si les droits ont chang� en cours de session
        EXCEPTIONEXTERNE(lbErrReqAdmin);
      *)
      IdUtilisateur:= FieldByName('Id_uti').AsInteger;  // m�morise l'utilisateur lui m�me
      OdtActif:= FieldByName('Id_odt').AsInteger;
      Close;
    end;
  finally
    QSession.Free;
  end;

  topHorloge:= Date+GetTime;
  with Command do
  begin
    CommandText:= format('update Session set ReqEnCours=1,TopRequete=%5.6f,Preferences=''%s'' where id_ses = %d',
     [topHorloge,codeAccesSession,numSessRecu],formatsADN);
    Execute;
  end;

{ v3.5.6: cr�ation d'un enregistrement d'avancement dans AvancementRequete }
  DerniereMajAvancement:= 0;   // v3.6.0: champ qui sera test� pour �viter des enregistrements r�p�t�s dans Avancement
                               // valeur 0 pour ne pas shunter le premier update
                               // (au cas o� le client teste tout de suite l'avancement pour avoir au moins un message)
  TauxAvancement:= 0;   // v3.6.0
  TauxAvancementMaxi:= 100;   // v3.6.0

  if IdRequete <> 0 then
    with Command do
    begin
      CommandText:= format(
       'insert AvancementRequete (Id_Ses,Id_Req,Pourcentage) values(%d,%d,0)',
       [NumSession,IdRequete]);
      Execute;
    end;

  Options:= 0;


  {D�tecter directement la cha�ne <param nom="longparam"><![CDATA[ apr�s le champ <infoSession> (aux espaces pr�s) }
  { Essaie de lire un attribut nom = "LongParam" }
  positFinBaliseParam:= posEx('>',peChRequete,ptrPremierParam);   // � bien faire avant de changer ptrPremierParam
  ptrPremierParam:= posEx(leftStr(defBalParam,pred(length(defBalParam))),peChRequete,ptrPremierParam);
  // on cherche la cha�ne '<param' (au cas o� il y aurait d'autres balises que <param nom = "...)
  // := 0 si ptrPremierParam
(*  if positFinBaliseParam > 0 then  *)
  if ptrPremierParam > 0 then   // v3.6.6j (1396)
  begin
    ValChampXmlSimple(defBalParam,midStr(peChRequete,ptrPremierParam,succ(positFinBaliseParam-ptrPremierParam))   // partie <param nom = "xxxxx">
     +FinBalise(defBalParam),aFournir,tabNom,tabVal,presente);
     // On ajoute </param> pour que ValChampXmlSimple puisse r�cup�rer les attributs <param nom = "xxxxx"></param>
     // (la valeur retourn�e sera vide)
     // On aurait pu appeler LitAttributs mais la proc. n'est pas d�clar�e dans l'interface de l'unit� ManipXml

    if (length(tabNom) > 0) and (tabNom[0] = defAttNom) then
    begin
      if uppercase(tabVal[0]) = uppercase(defAttLongParam) then
      begin
        chValParam:= trim(midStr(peChRequete,succ(positFinBaliseParam),
         length(peChRequete)-positFinBaliseParam-length(FinBalise(defBalParam))));
        if (leftStr(chValParam,length(cstDebutCData)) = cstDebutCData)
         and (rightStr(chValParam,length(cstFinCData)) = cstFinCData) then
        begin
          { isoler le contenu du param�tre long }
          chValParam:= midStr(chValParam,succ(length(cstDebutCData)),
           length(chValParam)-length(cstDebutCData)-length(cstFinCData));

          if LitRegistreADN(cstRegCheminsDAcces,cstRegR3Serveur,cheminR3Web) then
          begin
            lstLongParam:= TStringList.Create;
            try
              nomFicLongParam:= 'Tmp-Par'+CompleteAZero(NumSession,10)+'-'+CompleteAZero(IdRequete,10)+'.txt';

             if (cheminR3Web = '') or (cheminR3Web[length(cheminR3Web)]<>'\') then
                cheminR3Web:= cheminR3Web + '\';
              cheminR3Web:= cheminR3Web + 'TEMP\';

              if FileExists(cheminR3Web+nomficLongParam) then
                lstLongParam.LoadFromFile(cheminR3Web+nomficLongParam);

              if self is clsEnvoiLongParam then
              begin
                lstLongParam.Add(chValParam);
                lstLongParam.SaveToFile(cheminR3Web+nomficLongParam);
                traiterRequete:= false;
              end
              else
                { C'est la requ�te principale: les requ�tes pr�c�dentes �taient
                  uniquement destin�es � stocker les param�trs au pr�alable }
              begin
                if lstLongParam.Count = 0 then   // On n'a pas trouv� de param�tres pr�c�demment stock�s
                  EXCEPTIONINTERNE(defErr226);

                peChRequete:= lstLongParam.Strings[0];
                for iL:= 1 to pred(lstLongParam.Count) do
                  peChRequete:= peChRequete +lstLongParam.Strings[iL];
                peChRequete:= peChRequete +chValParam;   // ajout derni�re tranche de cha�ne de param�tres
                traiterRequete:= true;
              end
            finally
              lstLongParam.Free
            end;
          end
          else
            EXCEPTIONINTERNE(defErr300,lbErrCheminReponse);   // installation incorrecte
        end
          else
            EXCEPTIONINTERNE(defErr201,format(lbErrFormatLongParamIncorrect,[defAttLongParam]));

      end
      else
        traiterRequete:= true;

      if traiterRequete then
      begin

      { Stockage des param�tres dans les propri�t�s globales NomParam et ValParam }
        compteParam:= 0;
        repeat

          chValParam:= ValChampXml(defBalParam,peChRequete,tabNom,tabVal,presente,compteParam+1);
          if presente then
          begin
            inc(compteParam);
            SetLength(NomParam,compteParam);
            NomParam[pred(compteParam)]:= ValeurAttribut(defAttNom,tabNom,tabVal);
            // M�morise la valeur de l'attribut "nom" (de param�tre)
            if NomParam[pred(compteParam)] = '' then
              RAISE excReqIncor.Create(lbErrReqIncor);

            if upperCase(NomParam[pred(compteParam)]) = uppercase(defAttIgnorerAlerte) then
              IgnorerAlerte:= upperCase(chValParam) = 'O';
              
            SetLength(ValParam,compteParam);
            ValParam[pred(compteParam)]:= chValParam;
            // M�morise la valeur du param�tre lui-m�me qui est dans le champ XML
          end
          else
            BREAK
        until false;
      end;
    end
    else
      EXCEPTIONINTERNE(defErr201,lbErrAttributNomManquant);
  end
  else    // il n'y a aucun param�tre
    traiterRequete:= true;

  if traiterRequete then
  begin

    MajNomTablesPermanentes;

    if OdtActif > 0 then
    begin
      debutNom:= 'TMP'+CompleteAZero(NumSession,10);
      { 1) Changement de nom dans les 2 tableaux de noms de tables }
      for indCat:= low(TCategorie) to high(TCategorie) do
        if indCat <> eOrdretravaux then
          NomTableCompo[indCat]:= debutNom+NomTableCompo[indCat];
      for indAutre:= low(TAutreTable) to high(TAutreTable) do
        NomAutreTable[indAutre]:= debutNom+NomAutreTable[indAutre];
    end;

    EnregHistorique:= true;   // enregistrer les actions dans l'historique

  end;    // if traiterRequete
  SortieProc;
end;

{ ---------------------------------------------------------------------------- }
procedure clsReqGenerique.Finalise;
{ Remet l'indicateur "requ�te en cours" � 0 }
{ ---------------------------------------------------------------------------- }

begin
  EntreeProc(ClassName+'.Finalise');
  if ConnexionBD.Connected then
    with Command do
    begin
      CommandText:= format('update Session set ReqEnCours=0 where id_ses = %d',
      [NumSession]);
      Execute;
    end;

  { v3.5.6: les requ�tes longues poss�dent un IdRequete et une ligne correspondante dans la table AvancementRequete }
  if IdRequete <> 0 then
    with Command do
    begin
      CommandText:= format('delete from AvancementRequete where Id_Ses = %d and Id_req = %d',
       [NumSession,IdRequete]);
      Execute;
    end;
  SortieProc;
end;

{ ---------------------------------------------------------------------------- }
procedure clsReqGenerique.AnnuleTransactions;
{ Annule toutes les transactions en cours }
{ ---------------------------------------------------------------------------- }
begin
if ConnexionBD.InTransaction then
  ConnexionBD.RollbackTrans;
end;

{ ---------------------------------------------------------------------------- }
procedure clsReqGenerique.EntreeProc (pLibelProc: string);
{ empile dans la stringList PileAppels un nom de proc�dure ou un libell� sp�cial
(pour messages lors des exceptions)
 pLibelProc: libell� � empiler
 Appeler syst�matiquement EntreeProc au d�but de chaque proc�dure d'un objet
 d�riv� de clsReqGene }
{ ---------------------------------------------------------------------------- }

var
  ptChrono: ^TDateTime;

{$IFDEF MODECHRONO}
  oEltChrono: clsEltChrono;
{$ENDIF}

begin
  PileAppels.Add(pLibelProc);
{$IFDEF MODECHRONO}

  { M�morisation du top horloge de d�but de la pro�cdure }
  new(ptChrono);

  ptChrono^:= GetTime;
  PileChronos.Add(ptChrono);

  { Ajout �ventuel de l'entr�e dans le tableau TempsPasse }
  oEltChrono:= clsEltChrono.Create;
  oEltChrono.NomMethode:= pLibelProc;
  oEltChrono.Temps:= 0;
  oEltChrono.NbAppels:= 0;

  if not TempsPasse.Ajoute(oEltChrono) then
    oEltChrono.Free;
{$ENDIF}
end;

{ ---------------------------------------------------------------------------- }
procedure clsReqGenerique.SortieProc;
{ D�sempile la stringList PileAppels (lors de la sortie d'une proc�dure o� les
( exceptions sont g�r�es )
 Appeler syst�matiquement SortieProc en fin de chaque proc�dure non r�cursive  }
{ ---------------------------------------------------------------------------- }
var
  ptChrono: ^TDateTime;
  ptAllocMemoire: ^integer;
  topHorloge,tempsProc: TDateTime;
  indT: integer;
  oEltChrono: clsEltChrono;

begin
{$IFDEF MODECHRONO}
  with PileChronos do
    if Count > 0 then
    begin
      topHorloge:= GetTime;
      ptChrono:= Items[pred(Count)];
      { Rechercher le libell� de PileAppels.Items[pred(Count)] dans TempsPasse --> indice indT}
      oEltChrono:= clsEltChrono.Create;
      try
        oEltChrono.NomMethode:= PileAppels[pred(PileAppels.Count)];

        if TempsPasse.Trouve(oEltChrono,indT) then   // normalement toujours vrai
        begin
          tempsProc:= (topHorloge - ptChrono^) *86400;
          with TempsPasse[indT] as clsEltChrono do
          begin
            Temps:= Temps + tempsProc;
            inc(NbAppels);
          end;
        end
        else
          EXCEPTIONINTERNE(defErr142);  // oEltChrono non trouv� !

        dispose(ptChrono);
        delete(pred(count));

        if PileAppels.Count >= 2 then  // S'il y a une proc�dure appelante dans la pile
        begin
          { D�duire le temps qu'on vient de passer dans la proc�dure actuelle du temps pass� dans la proc�dure appelante }
          oEltChrono.NomMethode:= PileAppels[PileAppels.Count-2];
          if TempsPasse.Trouve(oEltChrono,indT) then   // normalement toujours vrai
            with TempsPasse[indT] as clsEltChrono do
              Temps:= Temps - tempsProc;
        end;

      finally
        oEltChrono.Free;
      end;
    end;


{$ENDIF}
  with PileAppels do
    if Count > 0 then
      Delete(pred(Count));        // Efface la derni�re cha�ne
end;

{ ---------------------------------------------------------------------------- }
function clsReqGenerique.AffichePileAppels: string;
{ ---------------------------------------------------------------------------- }

var i: integer;

begin
  if PileAppels.Count <> 0 then
    result:= lbProcedure + PileAppels[pred(PileAppels.Count)] + sautDeLigne
  else
    result:= '';
  for i:= PileAppels.Count - 2 downto 0 do
    result:= result + RendNonXml('<== ') + PileAppels[i] + sautDeLigne;
  PileAppels.Clear;
  result:= result + lbRequete + '=' + RendNonXml(CorpsRequete);
end;

{ --------------------------------------------------------------------- }
{ Renvoie une chaine de la forme <Message type = "peTypeMessage">peContenu</Message> }
function clsReqGenerique.RemplitMessage (peTypeMessage: tTypeMessage; peContenu: string): string;
{ peTypeMessage = indice de TypesMessages pointant sur le libell� � envoyer � DebutXml
{ ATTENTION : pas de codage par RendNonXml !
{ --------------------------------------------------------------------- }
var
  nomAttribMessage,valAttribMessage: tabAttrib;
  nbAttrib: integer;

begin
  EntreeProc(ClassName+'.RemplitMessage (1)');
  if DcnxPrevue <> 0 then
  begin   // On ajoute un attribut pour avertir de la d�connexion imminente
    nbAttrib:= 2;
    setLength(nomAttribMessage,2);
    setLength(valAttribMessage,2);
    nomAttribMessage[1]:= defAttDeconnexion;
    valAttribMessage[1]:=
     format(lbDeconnexionPrevue,[TimeToStr(DcnxPrevue-int(DcnxPrevue),FormatsADN)]);
    if MotifDeconnexion <> '' then
      valAttribMessage[1]:=
       valAttribMessage[1] + format(lbMotif,[RendNonXml(MotifDeconnexion)]);
  end
  else
  begin
    setLength(nomAttribMessage,1);
    setLength(valAttribMessage,1);
    nbAttrib:= 1;
  end;
  nomAttribMessage[0]:= defAttType;
  valAttribMessage[0]:= TypesMessage[peTypeMessage];

  if ModifInhibee then
  begin   { Le client pourra ainsi en d�but ou r�veil de session afficher le message "Passage en lecture seule" }
    inc(nbAttrib);
    setLength(nomAttribMessage,nbAttrib);
    setLength(valAttribMessage,nbAttrib);
    nomAttribMessage[pred(nbAttrib)]:= defAttDroit;
    valAttribMessage[pred(nbAttrib)]:= cstProfilModifInhibee;
  end;
  result:= RemplitBalise(defBalMessage, peContenu, nomAttribMessage, valAttribMessage, false);
  SortieProc;
end;

{ --------------------------------------------------------------------- }
{ Variante de RemplitMessage avec attributs suppl�mentaires pass�s en param�tre }
{ Renvoie une chaine de la forme
 <Message type = "peTypeMessage" attrib1 = "val1" attrib2 = "val2" ... >peContenu</Message> }
function clsReqGenerique.RemplitMessage (
 peTypeMessage: tTypeMessage;  // libell� � envoyer � DebutXml
 peContenu: string;
 peNomAttrib: array of string;   // tableau des noms d'attributs suppl�mentaires
 peValAttrib: array of string)   // tableau des valeurs d'attributs suppl�mentaires
 : string;
{ peTypeMessage =
{ ATTENTION : pas de codage par RendNonXml !
{ --------------------------------------------------------------------- }
var
  nomAttribMessage,valAttribMessage: tabAttrib;
  ia,nbAttribSuppl: integer;

begin
  EntreeProc(ClassName+'.RemplitMessage (2)');
  if DcnxPrevue = 0 then
    nbAttribSuppl:= 1
  else
    nbAttribSuppl:= 2;
  setLength(nomAttribMessage,nbAttribSuppl+length(peNomAttrib));
  setLength(valAttribMessage,nbAttribSuppl+length(peValAttrib));
  nomAttribMessage[0]:= defAttType;
  valAttribMessage[0]:= TypesMessage[peTypeMessage];

  if DcnxPrevue <> 0 then
  begin   // On ajoute un attribut pour avertir de la d�connexion imminente
    nomAttribMessage[1]:= defAttDeconnexion;
    valAttribMessage[1]:=
     format(lbDeconnexionPrevue,[TimeToStr(DcnxPrevue-int(DcnxPrevue),FormatsADN)]);
    if MotifDeconnexion <> '' then
      valAttribMessage[1]:=
       valAttribMessage[1] + format(lbMotif,[RendNonXml(MotifDeconnexion)]);
  end;

  { Copie des �l�ments de peNomAttrib et peValAttrib dans nomAttribMessage et ValAttribMessage }
  for ia:= 0 to high(peNomAttrib) do
  begin
    nomAttribMessage[ia+nbAttribSuppl]:= peNomAttrib[ia];
    valAttribMessage[ia+nbAttribSuppl]:= peValAttrib[ia];
  end;

  result:= RemplitBalise(defBalMessage, peContenu, nomAttribMessage, valAttribMessage, false);
  SortieProc;
end;

{ ---------------------------------------------------------------------------- }
function clsReqGenerique.RecupereIdent: integer;
{ renvoie l'identifiant automatiquement g�n�r� par la derni�re requ�te de cr�ation }
{ ---------------------------------------------------------------------------- }

begin
  EntreeProc(ClassName+'.RecupereIdent');
  with Query do
  begin
    SQL.Text:='Select SCOPE_IDENTITY() [ident]';
    open;

    result:= FieldByName('Ident').AsInteger;
    close;
  end;
  SortieProc;
end;

{ ---------------------------------------------------------------------------- }
procedure clsReqGenerique.SQLRecupereIdent(
 var peTexteSQL: TStringList; // contient le script SQL en cours de constitution
 peNomVarSQL: string;  // nom de la variable SQL utilis�e
 peDeclarationVar: boolean = true);   // indicateur ajout d�claration de (peNomVarSQL) dans le script
{ renvoie le script SQL donnant le dernier identifiant automatiquement g�n�r�  }
{ ---------------------------------------------------------------------------- }

begin
  EntreeProc(ClassName+'.SQLRecupereIdent');
  with peTexteSQL do
  begin
    if peDeclarationVar then
      Add('declare @'+peNomVarSQL+' integer');
    // Sinon la d�claration doit avoir �t� mise dans le script global par l'appelant
    Add('set @'+peNomVarSQL+' = scope_identity()');
  end;
  SortieProc;
end;

{ ---------------------------------------------------------------------------- }
function clsReqGenerique.ValeurParam(peNomParam: string): string;
{ Donne la valeur d'un param�tre transmis � la requ�te et stock� dans
        (NomParam,ValParam) }
{ Param�tre :
  peNomParam = nom du param�tre
{ ---------------------------------------------------------------------------- }

begin
  EntreeProc(ClassName+'.ValeurParam');
  result:= ValeurAttribut(peNomParam,NomParam,ValParam);
  SortieProc
end;

{ ---------------------------------------------------------------------------- }
function clsReqGenerique.ValeurParam(peNomParam: string;
 peValeurDefaut: integer)   // valeur par d�faut si param�tre vide
 : integer;
{ Similaire � la pr�c�dente, mais donne directement un r�sultat de type entier }
{ ---------------------------------------------------------------------------- }
var paramChaine: string;

begin
  EntreeProc(ClassName+'.ValeurParam');
  paramChaine:= ValeurAttribut(peNomParam,NomParam,ValParam);
  if paramChaine = '' then
    result:= peValeurDefaut
  else
    if not TryStrToInt(paramChaine,result) then
      ExceptionInterne(defErr201,'('+peNomParam+')');  // param�tres de requ�te incorrects
  SortieProc
end;


{ ----------------------------------------------------------------------------------------- }
function clsReqGenerique.RemplitItem (peNom: string; peCategorie: char; peNum: integer = -1;
 peInfo: integer = 0): string;
{ param�tres :
*   - peNom
*   - peCategorie  = eLieuSimple,eLocalTechnique,eCable, etc.
*   - peNum (num�rique)
*   - peInfo (fac) : information additionnelle (caract�re, num�rique ou date)
{ ----------------------------------------------------------------------------------------- }

begin
  EntreeProc(className+'.RemplitItem');
  result:= RemplitBalise(defBalNom, trim(peNom), true);    // avec codage "non XML"
  result:= result + RemplitBalise(defBalCategorie, peCategorie);
  if peNum <> -1 then
    result:= result + RemplitBalise(defBalNum, IntToStr(peNum));
  if peInfo <> 0 then
    result:= result + RemplitBalise(defBalInfo, IntToStr(peInfo));
  SortieProc;
end;

{ ----------------------------------------------------------------------------------------- }
{ Variante de la pr�c�dente avec param�tres obligatoires et peInfo sous forme de cha�ne }
function clsReqGenerique.RemplitItem (peNom: string; peCategorie: char; peNum: integer;
 peInfo: string): string;
{ param�tres :
*   - peNom
*   - peCategorie  = eLieuSimple,eLocalTechnique,eCable, etc.
*   - peNum (num�rique)
*   - peInfo (fac) : information additionnelle (caract�re, num�rique ou date)
{ ----------------------------------------------------------------------------------------- }

begin
  EntreeProc(className+'.RemplitItem');
  result:= RemplitBalise(defBalNom, trim(peNom), true);    // avec codage "non XML"
  result:= result + RemplitBalise(defBalCategorie, peCategorie);
  if peNum <> -1 then
    result:= result + RemplitBalise(defBalNum, IntToStr(peNum));
  if peInfo <> '' then
    result:= result + RemplitBalise(defBalInfo, peInfo, true);    // avec codage "non XML"
  SortieProc;
end;

{ ----------------------------------------------------------------------------------------- }
function clsReqGenerique.DoubleQuotes(peChaine: string;
 peCarDelim: char = '''';  // peCarDelim : caract�re servant de d�limiteur de cha�ne ajout� v3.4.3
 peEncadrerChaine: boolean = false   // true s'il faut ajouter peCarDelim au d�but et � la fin de la cha�ne v3.4.3
 ): string;  
{ Double les "'" dans une cha�ne pour pouvoir la soumettre comme param�tre dans une
        requ�te SQL }
{ Param�tre peChaine = cha�ne d'origine }
{ Renvoie la cha�ne r�sultante }
{ ----------------------------------------------------------------------------------------- }
var posit,debutRech,longueur: integer;

begin
  EntreeProc(className+'.DoubleQuotes');
  debutRech:= 0;
  longueur:= length(peChaine);  // minuscule optimisation
  result:= peChaine;
  repeat
    posit:= pos(peCarDelim,midStr(result,succ(debutRech),longueur));
    if posit = 0 then
      BREAK;
    debutRech:= succ(debutRech + posit);  // succ() car �a a allong� la cha�ne result
    result:= stuffString(result,debutRech,0,peCarDelim);    // Ajoute un ' � c�t� de l'autre
  until false;
  if peEncadrerChaine then
    result:= peCarDelim+result+peCarDelim;
  SortieProc;
end;

{ ----------------------------------------------------------------------------------------- }
function clsReqGenerique.TraiteException(peExc: exception): string;
{ Pr�pare le message � renvoyer au client en fonction de l'exception peExc
  (l'exception a �t� d�clench�e par programme ou par erreur d'ex�cution) }
{ ----------------------------------------------------------------------------------------- }

begin
  if (peExc is excIncohBD) or (peExc is excIncohProg) or (peExc is excIncohClient) then
    result:= RemplitMessage(eMessErreurProgramme,
     peExc.Message+sautDeLigne+AffichePileAppels)
  else
    begin
    if peExc is excEnVeille then
      result:= RemplitMessage(eMessMdpRequis,peExc.Message)
    else
      if peExc is excArretTraitement then   // Annulation traitement car op�ration impossible
      begin
        while PileAppels.Count > 0 do
          SortieProc;   // Utile pour le chronom�trage

        result:= RemplitMessage(eMessErreurUtilisateur,lbErrExterne+peExc.Message);
      end
      else
        if peExc is excDephasage then
        begin
          result:= RemplitMessage(eMessDephasage,peExc.Message);
          while PileAppels.Count > 0 do
            SortieProc;   // Utile pour le chronom�trage
        end
        else
          if peExc is excDemandeConfirm then
            result:= RemplitMessage(eMessDialogue,peExc.Message)
          else
            if peExc is excDialogueSpecial then 
              result:= RemplitMessage(eMessDialogueSpecial,peExc.Message)
            else
              if peExc is excArretTotal then
                result:= RemplitMessage(eMessStop,peExc.Message)
              else                       // Erreur d'ex�cution
              begin
                if (pos(lbErrSQLServerInterblocageFR,peExc.Message) > 0)  // 3.6.6c (1331)
                 or (pos(lbErrSQLServerInterblocageEN,peExc.Message) > 0) then  // Si SQL Server anglais
                  result:= RemplitMessage(eMessStop,lbErrTransactionRejetee)
                else
                  result:= RemplitMessage(eMessErreurProgramme,
                   lbErrExec+peExc.Message+sautDeLigne+AffichePileAppels)
              end;
    end;
end;


{ -----------------doc dans LigneeItemLieux.txt ---------------------------------------- }
procedure clsReqGenerique.LigneeItemLieux(peCateg: TCategorie; peIdObj: integer;
     var psRangDansArbre: string; var psLignee: string; var psDroit: TDroit);
{ Met dans les cha�nes psRangDansArbre et psLignee la lign�e (tous les ascendants)
 d'un item de l'arbre des lieux.


Param�tres d'entr�e :
*   peCateg =  cat�gorie de l'objet de r�f�rence
* peIdObj  = n� interne de l'objet de r�f�rence

Param�tres de sortie :
*       psRangDansArbre :
 C'est une suite de s�quences, en commen�ant par
 le niveau site (la vue g�n�rale est exclue car unique).
        <chr(15)><N� d'ordre de lieu cadr� � droite sur 5 car.>
    ou  <chr(16)><Nom de local technique>
    ou  <chr(17)><Nom de groupe d'�quipement>
    ou  <chr(18)><Nom d'�quipement>
    ou  <chr(19)><Nom de terminaison>
 Les chr(xx) �tant inf�rieurs � tout caract�re alphab�tique, l'ordre alphab�tique
 de la cha�ne d�termine le classement dans l'arbre, m�me si les noms sont de longueur
 variable.

*       psLignee = cha�ne Xml conforme � la description Xml d'arbre R3Web
 Elle contient les infos de tous les ascendants de l'objet jusqu'au niveau site
 (pas la vue g�n�rale)
 Elle contient en outre un attribut suppl�mentaire donnant l'�ventuel droit du lieu
 ( ce n'est pas le droit induit) (sinon 0)
* psDroit = droit d'acc�s de l'item, d�duit des ses anc�tres s'il n'a pas de droit sp�cifique
{ -------------------------------------------------------------------------------------- }

var
  QLignee: TADOQuery;
  categLieu1,separateur: char;
  chDiscri,nomCour: string;
  topNiveau,indNiveau: smallint;
  prochainPere: integer;
  droitLu: TDroit;

begin
  EntreeProc(ClassName+'.LigneeItemLieux');
  psRangDansArbre:= '';
  psLignee:= '';
  try
    QLignee:= CreeADOQuery();
    with QLignee do
    begin
      if peCateg = eLieuSimple then
      begin
        SQL.Add(format('select L.Nom, L.Id_type, T.NumNiveau, L.Classement, L.IdPere, D.DroitLieu from %s L',
         [NomTablecompo[eLieuSimple]]));
        SQL.Add('join Type T on T.Id_type = L.Id_type');
        SQL.Add(format('left join Droit_lieu D on D.Id_prof = %d and D.Id_lieu=L.Id_lieu',
         [IdProfil]));
        SQL.Add(format('where L.Id_lieu = %d',[peIdObj]));
        Open;
        if Eof then
          EXCEPTIONEXTERNE(format(lbErrCompAbsent,[InfoCateg[peCateg].Libelle,intToStr(peIdObj)]));
        str(FieldByName('Classement').AsInteger :5, chDiscri);
        psDroit:= FieldByName('DroitLieu').AsInteger;
        // peut �tre � 0 si NULL (pas de droit explicite dans Droit_lieu)

        psLignee:= RemplitBalise(defBalItem, RemplitItem(
         FieldByName('Nom').AsString, InfoCateg[peCateg].Code, peIdObj,
         FieldByName('Classement').AsInteger),
         [defAttSelection,defAttDroit,defAttType],['O',intToStr(psDroit),FieldByName('Id_type').AsString]);
        { Le s�parateur sert � classer les �l�ments d'arbre par simple comparaison
          de la cha�ne caract�ristique psRangDansArbre }
        psRangDansArbre:= cstPrefixeLieuSimple + chDiscri;
        topNiveau:= FieldByName('NumNiveau').AsInteger;
        prochainPere:= FieldByName('IdPere').AsInteger;
      end

      else
      begin
        case peCateg of
          eEquipement,eTerminaison:
          begin
            SQL.Add('select E.Nom as NomObj, E.Id_type as typeEqt,');
            SQL.Add('L1.Nom as Nom1, L2.Nom as Nom2, L3.Nom as Nom3,');
            SQL.Add('L1.Id_lieu as id_Lieu1, L2.Id_lieu as id_Lieu2, L3.Id_lieu as id_Lieu3,');
            SQL.Add('TL1.Categorie as CategPere, TL2.NumNiveau as NN2, TL3.NumNiveau as NN3,');
            SQL.Add('L3.IdPere as IdPere3,');
            SQL.Add('D1.DroitLieu as DL1, D2.DroitLieu as DL2, D3.DroitLieu as DL3,');
            SQL.Add('L1.Classement as Cls1, L2.Classement as Cls2 ,L3.Classement as Cls3');
            SQL.Add(format('from %s E join %s L1 on L1.Id_lieu = E.Id_lieu',
             [NomTableCompo[eEquipement],NomTableCompo[eGroupe]]));
            SQL.Add('join Type TL1 on TL1.Id_type = L1.Id_type');
            SQL.Add(format('join %s L2 on L2.Id_lieu = L1.IdPere',
             [NomTableCompo[eLocalTechnique]]));
            SQL.Add('join Type TL2 on TL2.Id_type = L2.Id_type');
            SQL.Add(format('left join %s L3 on L3.Id_lieu = L2.IdPere',
             [NomTableCompo[eLieuSimple]]));
            SQL.Add('left join Type TL3 on TL3.Id_type = L3.Id_type');
            // NB: si l'�quipement n'appartient pas � un GE, cela fait remonter
            // jusqu'au p�re du LS contenant le LT de l'eqt.
            // Si c'est une terminaison d�finie directement dans un site, L3 est NULL
            // gr�ce au left join. Ca ne marche que parce que
            // NomTableCompo[eGroupe]=NomTableCompo[eLocalTechnique]=NomTableCompo[eLieuSimple]
            SQL.Add(format('left join Droit_lieu D1 on D1.Id_prof = %d and D1.Id_lieu=L1.Id_lieu',
             [IdProfil]));
            SQL.Add(format('left join Droit_lieu D2 on D2.Id_prof = %d and D2.Id_lieu=L2.Id_lieu',
             [IdProfil]));
            SQL.Add(format('left join Droit_lieu D3 on D3.Id_prof = %d and D3.Id_lieu=L3.Id_lieu',
             [IdProfil]));
            SQL.Add(format('where E.Id_eqt = %d',[peIdObj]));
            Open;
            if Eof then
              EXCEPTIONEXTERNE(format(lbErrCompAbsent,[InfoCateg[peCateg].Libelle,intToStr(peIdObj)]));

            { Remonter les ascendants de l'�quipement jusqu'au grand-p�re ou � l'arri�re grand-p�re
              et construire les listes }
            if peCateg = eTerminaison then
              psRangDansArbre:= cstPrefixeTerminaison
            else
              psRangDansArbre:= cstPrefixeEquipement;
            psRangDansArbre:= psRangDansArbre + FieldByName('NomObj').AsString;
            psLignee:= RemplitBalise(defBalItem, RemplitItem(
             FieldByName('NomObj').AsString, InfoCateg[peCateg].Code, peIdObj),
             [defAttSelection,defAttType],['O',FieldByName('TypeEqt').AsString]);
            categLieu1:= FieldByName('CategPere').AsString[1];
            // categLieu1 = cat�gorie du 1er lieu rencontr� en remontant l'arbre
          end;   //       cas eEquipement,eTerminaison

          eGroupe,eLocalTechnique:
          begin   { On est s�r qu'on peut remonter jusqu'au grand-p�re qui ne peut
                    pas �tre plus haut que le niveau vue g�n�rale dans le cas d'un LT }
            SQL.Add('select L1.Nom as Nom1, L1.Id_type, L2.Nom as Nom2, L3.Nom as Nom3,');
            SQL.Add('L1.Id_lieu as id_Lieu1, L2.Id_lieu as id_Lieu2, L3.Id_lieu as id_Lieu3,');
            SQL.Add('L3.IdPere as IdPere3,');
            SQL.Add('D1.DroitLieu as DL1, D2.DroitLieu as DL2, D3.DroitLieu as DL3,');
            SQL.Add('TL2.Categorie as CatL2, TL2.NumNiveau as NN2, TL3.NumNiveau as NN3,');
            SQL.Add('L1.Classement as Cls1, L2.Classement as Cls2 ,L3.Classement as Cls3');
            SQL.Add(format('from %s L1 join %s L2 on L2.Id_Lieu = L1.IdPere',
             [NomTableCompo[eGroupe],NomTableCompo[eLocalTechnique]]));
            SQL.Add('join Type TL2 on TL2.Id_type = L2.Id_type');  // ici le niveau max est Site
            SQL.Add(format('join %s L3 on L3.Id_Lieu = L2.IdPere',
             [NomTableCompo[eLieuSimple]]));
            SQL.Add('join Type TL3 on TL3.Id_type = L3.Id_type');  // ici le niveau max est vue g�n�rale
            SQL.Add(format('left join Droit_lieu D1 on D1.Id_prof = %d and D1.Id_lieu=L1.Id_lieu',
             [IdProfil]));
            SQL.Add(format('left join Droit_lieu D2 on D2.Id_prof = %d and D2.Id_lieu=L2.Id_lieu',
             [IdProfil]));
            SQL.Add(format('left join Droit_lieu D3 on D3.Id_prof = %d and D3.Id_lieu=L3.Id_lieu',
             [IdProfil]));
            SQL.Add(format('where L1.Id_lieu = %d',[peIdObj]));
            Open;
            if Eof then
              EXCEPTIONEXTERNE(format(lbErrCompAbsent,[InfoCateg[peCateg].Libelle,intToStr(peIdObj)]));
            { Modif 3.5.3 pour qu'on puisse transmettre indiff�remment peCateg � eGroupe ou � eLocalTechnique
              pour un m�me lieu technique dont on ne conna�t pas pr�cis�ment la cat�gorie }
            if FieldByName('CatL2').AsString = InfoCateg[eLocalTechnique].Code then
              peCateg:= eGroupe    // change �ventuellement peCateg si on l'avait transmis � eLocalTechnique dans l'ignorance de sa vraie cat�gorie
            else
              peCateg:= eLocalTechnique;  // change �ventuellement peCateg ...
            categLieu1:= InfoCateg[peCateg].Code;
          end;
        end;  // case peCateg of

        { Analyse du lieu de plus bas niveau : groupe, LT ou LS }
        psDroit:= FieldByName('DL1').AsInteger;
        // peut �tre � 0 si NULL (pas de droit explicite dans Droit_lieu)
        nomCour:= FieldByName('Nom1').AsString;
        if categLieu1 = InfoCateg[eLieuSimple].Code then
        begin
          separateur:= cstPrefixeLieuSimple;
          str(FieldByName('Cls1').AsInteger :5, chDiscri)
        end
        else
        begin
          chDiscri:= nomCour;
          if categLieu1 = InfoCateg[eLocalTechnique].Code then
            separateur:= cstPrefixeLocalTechnique
          else   // le p�re est un GE
            separateur:= cstPrefixeGroupe;
        end;
        if peCateg in [eLocalTechnique,eGroupe] then
          psLignee:= RemplitBalise(defBalItem, RemplitItem(
           nomCour, InfoCateg[peCateg].Code, peIdObj),
           [defAttSelection,defAttDroit,defAttType],['O',intToStr(psDroit),FieldByName('Id_type').AsString])
        else
          psLignee:= RemplitBalise(defBalItem, RemplitItem(
           nomCour, categLieu1, FieldByName('Id_Lieu1').AsInteger,
           FieldByName('Cls1').AsInteger)  // Cls1 doit �tre � NULL donc 0 pour un LT ou un GE
           + RemplitBalise (defBalFils, psLignee),[defAttDroit],[intToStr(psDroit)] );
        { Le s�parateur sert � classer les �l�ments d'arbre par simple comparaison
          de la cha�ne caract�ristique psRangDansArbre }
        psRangDansArbre:= separateur + chDiscri+ psRangDansArbre;

        { Analyse du p�re du lieu de plus bas niveau (L2) : LT ou LS }
        if (peCateg = eTerminaison) and (categLieu1 = InfoCateg[eLieuSimple].Code)
         and (FieldByName('NN2').AsInteger = 0)  then
        begin
          topNiveau:= 1;  // cas d'une terminaison d�finie juste en dessous d'un site
          if psDroit = 0 then
            psDroit:= FieldByName('DL2').AsInteger;  // ajout 3.3.1a : droit sur la vue g�n�rale
        end
        else
        begin
          nomCour:= FieldByName('Nom2').AsString;
          if separateur  = cstPrefixeGroupe then
          begin    // Le grand-p�re ne peut �tre qu'un local technique
            separateur:= cstPrefixeLocalTechnique;
            categLieu1:= InfoCateg[eLocalTechnique].Code;
            chDiscri:= nomCour
          end
          else   // Le grand-p�re est un lieu simple (ce peut �tre la vue g�n�rale !)
          begin
            separateur:= cstPrefixeLieuSimple;
            categLieu1:= InfoCateg[eLieuSimple].Code;
            str(FieldByName('Cls2').AsInteger :5, chDiscri);
          end;
          droitLu:= FieldByName('DL2').AsInteger;
          psLignee:= RemplitBalise(defBalItem, RemplitItem(
           nomCour, categLieu1, FieldByName('Id_Lieu2').AsInteger,
           FieldByName('Cls2').AsInteger)  // Cls2 doit �tre � NULL donc 0 pour un LT
           + RemplitBalise (defBalFils, psLignee) ,[defAttDroit],[intToStr(droitLu)]);
          { Le s�parateur sert � classer les �l�ments d'arbre par simple comparaison
            de la cha�ne caract�ristique psRangDansArbre }
          psRangDansArbre:= separateur + chDiscri+ psRangDansArbre;
          if psDroit = 0 then  // si pas encore lu un droit explicite
            psDroit:= droitLu;
            // peut �tre � 0 si NULL (pas de droit explicite dans Droit_lieu)

          { analyse du grand-p�re du lieu de plus bas niveau (L3)
            (il peut �tre NULL ou = vue g�n�rale)}
          topNiveau:= FieldByName('NN3').AsInteger;
          if topNiveau = 0 then
          begin
            if psDroit = 0 then
              psDroit:= FieldByName('DL3').AsInteger;  // = 0 si NN3 est NULL (donc DL3 aussi)
          end
          else   // topNiveau >= 1
          begin  { ce ne peut �tre qu'un lieu simple }
            prochainPere:= FieldByName('IdPere3').AsInteger;
            str(FieldByName('Cls3').AsInteger :5, chDiscri);
            droitLu:= FieldByName('DL3').AsInteger;
            psLignee:= RemplitBalise(defBalItem, RemplitItem(
             FieldByName('Nom3').AsString, InfoCateg[eLieuSimple].Code,
             FieldByName('Id_Lieu3').AsInteger, FieldByName('Cls3').AsInteger)
             + RemplitBalise (defBalFils, psLignee), [defAttDroit],[intToStr(droitLu)]);
            psRangDansArbre:= cstPrefixeLieuSimple + chDiscri+ psRangDansArbre;
            if psDroit = 0 then  // si pas encore lu un droit explicite
              psDroit:= droitLu;
              // peut �tre � 0 si NULL (pas de droit explicite dans Droit_lieu)
          end
        end;  // if FieldByName('NN2').AsInteger = 0 ... else ...
      end;  // if peCateg = eLieuSimple ... else ...

      { Si on ne l'a pas d�j� atteint, remont�e jusqu'au niveau 1 (site) }
      { La requ�te est variable en fonction du nombre de niveaux � remonter }
      if topNiveau > 1 then
      begin
        SQL.Clear;
        SQL.Add('select L1.Nom as Nom1, L1.Id_lieu as id_Lieu1, L1.Classement as Cls1');
        SQL.Add(', D1.DroitLieu as DL1');
        { On fait autant de jointures que de niveaux restants pour atteindre le niveau vue g�n�rale}
        for indNiveau:= 2 to pred(topNiveau) do
        begin
          SQL.Add(format(
           ', L%0:d.Nom as Nom%0:d, L%0:d.Id_lieu as Id_Lieu%0:d, L%0:d.Classement as Cls%0:d',
           [indNiveau]));
          SQL.Add(format(', D%0:d.DroitLieu as DL%0:d',[indNiveau]));
        end;
        { Droit sur la vue g�n�rale }
        SQL.Add(format(',D%0:d.DroitLieu as DL%0:d',[topNiveau]));
        SQL.Add(format('from %s L1',[NomTableCompo[eLieuSimple]]));
        SQL.Add(format('left join Droit_lieu D1 on D1.Id_prof = %d and D1.Id_lieu=L1.Id_lieu',
         [IdProfil]));
        for indNiveau:= 2 to pred(topNiveau) do
        begin
          SQL.Add(format('join %0:s L%1:d on L%1:d.Id_lieu = L%2:d.IdPere' ,
           [NomTableCompo[eLieuSimple],indNiveau,pred(indNiveau)]));
          SQL.Add(format('left join Droit_lieu D%0:d on D%0:d.Id_prof = %1:d and D%0:d.Id_lieu=L%0:d.Id_lieu',
           [indNiveau,IdProfil]));
        end;
        { Droit sur la vue g�n�rale: son num�ro est le IdPere du lieu de plus haut niveau lu }
        SQL.Add(format('left join Droit_lieu D%0:d on D%0:d.Id_prof = %1:d and D%0:d.Id_lieu=L%2:d.IdPere',
         [topNiveau,IdProfil,pred(topNiveau)]));
        SQL.Add(format('where L1.Id_lieu = %d',[prochainPere]));
        Open;
        if Eof then
          EXCEPTIONEXTERNE(format(lbErrCompAbsent,[InfoCateg[peCateg].Libelle,intToStr(peIdObj)]));

        { Compl�te la cha�ne caract�ristique et la lign�e Xml }
        for indNiveau:= 1 to pred(topNiveau) do
        begin
          str(fieldByName('Cls'+IntToStr(indNiveau)).AsInteger :5, chDiscri);
          psRangDansArbre:= cstPrefixeLieuSimple
           + chDiscri + psRangDansArbre;
          droitLu:= FieldByName('DL'+IntToStr(indNiveau)).AsInteger;
          psLignee:= RemplitBalise(defBalItem, RemplitItem(
           FieldByName('Nom'+IntToStr(indNiveau)).AsString, InfoCateg[eLieuSimple].Code,
           FieldByName('Id_lieu'+IntToStr(indNiveau)).AsInteger,
           FieldByName('Cls'+IntToStr(indNiveau)).AsInteger)
           + RemplitBalise (defBalFils, psLignee), [defAttDroit],[intToStr(droitLu)] );
          if (psDroit = 0) or (droitLu = cstDroitAucunAcces) then
            // si pas encore lu un droit explicite ou aucun acc�s (au niveau site)
            // (le droit "aucun acc�s" prime sur tout autre droit mais en principe il ne peut pas y avoir
            // de droits particuliers pour des lieux d'un site en "aucun acc�s"
            psDroit:= droitLu;
        end;
        { Si n�cessaire, examen du droit de la vue g�n�rale }
        if psDroit = 0 then
          psDroit:= FieldByName('DL'+IntToStr(topNiveau)).AsInteger;
      end    // if topNiveau > 1
      else
        if (topNiveau = 1) and (psDroit = 0) and not Administrateur then
          // On a atteint le niveau site mais il faut lire les droits sur la vue g�n�rale
        begin
          SQL.Clear;
          SQL.Add(format('select DroitLieu from Droit_lieu where Id_lieu = %d and Id_prof =%d',
           [prochainPere,IdProfil]));
          Open;
          psDroit:= fieldByName('DroitLieu').AsInteger;
        end;
    end;    // with QLignee

    if (ModifInhibee or ConsultationSeule) and (psDroit <= cstDroitModif) then
      psDroit:= cstDroitLectureSeule   // session d�grad�e en lecture seule
    else
      if psDroit = 0 then   // jamais trouv� de droit explicite
        psDroit:= cstDroitModif;  // pas de droit <=> tous les droits

  finally
    QLignee.Free;
  end;
  SortieProc;
end;

{ -------------------------------------------------------------------------------------- }
procedure clsReqGenerique.OrdreEtDroitLieu(
 peCateg: TCategorie;    // eLieuSimple ou eLocalTechnique ou eGroupe
 peIdObj: integer;       // n� du lieu
 var psOrdre: string;    // maj avec le champ Ordre (uniquement si peCateg = eLieuSimple)
 var psDroit: TDroit);   // maj avec le droit d'acc�s effectif du lieu
{ v3.5.2b M�thode analogue � LigneeItemLieux sauf que le lieu peut �tre la vue g�n�rale elle-m�me}
{ -------------------------------------------------------------------------------------- }

var
  reqLieu: TADOQuery;
  prochainPere,topNiveau,indNiveau,droitLu: integer;

begin
  EntreeProc(ClassName+'.DroitEtOrdreLieu');
  reqLieu:= CreeADOQuery();
  try
    with reqLieu do
    begin
      if peCateg = eLieuSimple then
      begin
        SQL.Add(format('select L.Nom, L.Id_type, T.NumNiveau, L.Ordre, L.IdPere, D.DroitLieu from %s L',
         [NomTablecompo[eLieuSimple]]));
        SQL.Add('join Type T on T.Id_type = L.Id_type');
        SQL.Add(format('left join Droit_lieu D on D.Id_prof = %d and D.Id_lieu=L.Id_lieu',
         [IdProfil]));
        SQL.Add(format('where L.Id_lieu = %d',[peIdObj]));
        Open;

        if Eof then
          EXCEPTIONEXTERNE(format(lbErrCompAbsent,[InfoCateg[peCateg].Libelle,intToStr(peIdObj)]));

        topNiveau:= FieldByName('NumNiveau').AsInteger;
        if topNiveau = 0 then
          prochainPere:= peIdObj   // c'est la VG - prochainPere sera utilis� pour chercher son droit d'acc�s
        else
          prochainPere:= FieldByName('IdPere').AsInteger;
        psOrdre:= FieldByName('Ordre').AsString
      end
      else
      begin   { On est s�r qu'on peut remonter jusqu'au grand-p�re qui ne peut
                pas �tre plus haut que le niveau vue g�n�rale dans le cas d'un LT }
        SQL.Add('select L1.Id_type,');
        SQL.Add('L1.Id_lieu as id_Lieu1, L2.Id_lieu as id_Lieu2, L3.Id_lieu as id_Lieu3,');
        SQL.Add('L3.IdPere as IdPere3,');
        SQL.Add('D1.DroitLieu as DL1, D2.DroitLieu as DL2, D3.DroitLieu as DL3,');
        SQL.Add('TL2.NumNiveau as NN2, TL3.NumNiveau as NN3');
        SQL.Add(format('from %s L1 join %s L2 on L2.Id_Lieu = L1.IdPere',
         [NomTableCompo[eGroupe],NomTableCompo[eLocalTechnique]]));
        SQL.Add('join Type TL2 on TL2.Id_type = L2.Id_type');  // ici le niveau max est Site
        SQL.Add(format('join %s L3 on L3.Id_Lieu = L2.IdPere',
         [NomTableCompo[eLieuSimple]]));
        SQL.Add('join Type TL3 on TL3.Id_type = L3.Id_type');  // ici le niveau max est vue g�n�rale
        SQL.Add(format('left join Droit_lieu D1 on D1.Id_prof = %d and D1.Id_lieu=L1.Id_lieu',
         [IdProfil]));
        SQL.Add(format('left join Droit_lieu D2 on D2.Id_prof = %d and D2.Id_lieu=L2.Id_lieu',
         [IdProfil]));
        SQL.Add(format('left join Droit_lieu D3 on D3.Id_prof = %d and D3.Id_lieu=L3.Id_lieu',
         [IdProfil]));
        SQL.Add(format('where L1.Id_lieu = %d',[peIdObj]));
        Open;

        if Eof then
          EXCEPTIONEXTERNE(format(lbErrCompAbsent,[InfoCateg[peCateg].Libelle,intToStr(peIdObj)]));
        { Analyse du lieu de plus bas niveau : groupe, LT ou LS }

        psDroit:= FieldByName('DL1').AsInteger;
        // peut �tre � 0 si NULL (pas de droit explicite dans Droit_lieu)

        { Analyse du p�re du lieu de plus bas niveau (L2) : LT ou LS }
        droitLu:= FieldByName('DL2').AsInteger;
        if psDroit = 0 then  // si pas encore lu un droit explicite
          psDroit:= droitLu;
          // peut �tre � 0 si NULL (pas de droit explicite dans Droit_lieu)

        { analyse du grand-p�re du lieu de plus bas niveau (L3)
          (il peut �tre NULL ou = vue g�n�rale)}
        topNiveau:= FieldByName('NN3').AsInteger;
        if topNiveau = 0 then
        begin
          if psDroit = 0 then
            psDroit:= FieldByName('DL3').AsInteger;  // = 0 si NN3 est NULL (donc DL3 aussi)
        end
        else   // topNiveau >= 1
        begin  { ce ne peut �tre qu'un lieu simple }
          prochainPere:= FieldByName('IdPere3').AsInteger;
          droitLu:= FieldByName('DL3').AsInteger;
          if psDroit = 0 then  // si pas encore lu un droit explicite
            psDroit:= droitLu;
            // peut �tre � 0 si NULL (pas de droit explicite dans Droit_lieu)
        end;
      end;

      { Si on ne l'a pas d�j� atteint, remont�e jusqu'au niveau 1 (site) }
      { La requ�te est variable en fonction du nombre de niveaux � remonter }
      if topNiveau > 1 then
      begin
        SQL.Clear;
        SQL.Add('select L1.Id_lieu as id_Lieu1, D1.DroitLieu as DL1');
        { On fait autant de jointures que de niveaux restants pour atteindre le niveau vue g�n�rale}
        for indNiveau:= 2 to pred(topNiveau) do
          SQL.Add(format(', L%0:d.Id_lieu as Id_Lieu%0:d, D%0:d.DroitLieu as DL%0:d',
           [indNiveau]));
        { Droit sur la vue g�n�rale }
        SQL.Add(format(',D%0:d.DroitLieu as DL%0:d',[topNiveau]));
        SQL.Add(format('from %s L1',[NomTableCompo[eLieuSimple]]));
        SQL.Add(format('left join Droit_lieu D1 on D1.Id_prof = %d and D1.Id_lieu=L1.Id_lieu',
         [IdProfil]));
        for indNiveau:= 2 to pred(topNiveau) do
        begin
          SQL.Add(format('join %0:s L%1:d on L%1:d.Id_lieu = L%2:d.IdPere' ,
           [NomTableCompo[eLieuSimple],indNiveau,pred(indNiveau)]));
          SQL.Add(format('left join Droit_lieu D%0:d on D%0:d.Id_prof = %1:d and D%0:d.Id_lieu=L%0:d.Id_lieu',
           [indNiveau,IdProfil]));
        end;
        { Droit sur la vue g�n�rale: son num�ro est le IdPere du lieu de plus haut niveau lu }
        SQL.Add(format('left join Droit_lieu D%0:d on D%0:d.Id_prof = %1:d and D%0:d.Id_lieu=L%2:d.IdPere',
         [topNiveau,IdProfil,pred(topNiveau)]));
        SQL.Add(format('where L1.Id_lieu = %d',[prochainPere]));
        Open;
        if Eof then
          EXCEPTIONEXTERNE(format(lbErrCompAbsent,[InfoCateg[peCateg].Libelle,intToStr(peIdObj)]));

        for indNiveau:= 1 to pred(topNiveau) do
        begin
          droitLu:= FieldByName('DL'+IntToStr(indNiveau)).AsInteger;
          if (psDroit = 0) or (droitLu = cstDroitAucunAcces) then
            // si pas encore lu un droit explicite ou aucun acc�s (au niveau site)
            // (le droit "aucun acc�s" prime sur tout autre droit mais en principe il ne peut pas y avoir
            // de droits particuliers pour des lieux d'un site en "aucun acc�s"
            psDroit:= droitLu;
        end;

        { Si n�cessaire, examen du droit de la vue g�n�rale }
        if psDroit = 0 then
          psDroit:= FieldByName('DL'+IntToStr(topNiveau)).AsInteger;
      end    // if topNiveau > 1
      else
        if (topNiveau = 0)   //  La racine de la branche � d�velopper est la vue g�n�rale elle-m�me
         or (topNiveau = 1) and (psDroit = 0) and not Administrateur then
          // On a atteint le niveau site mais il faut lire les droits sur la vue g�n�rale
        begin
          SQL.Clear;
          SQL.Add(format('select DroitLieu from Droit_lieu where Id_lieu = %d and Id_prof =%d',
           [prochainPere,IdProfil]));
          Open;
          psDroit:= fieldByName('DroitLieu').AsInteger;
        end

    end;    // with reqLieu

    if (ModifInhibee or ConsultationSeule) and (psDroit <= cstDroitModif) then
      psDroit:= cstDroitLectureSeule   // session d�grad�e en lecture seule
    else
      if psDroit = 0 then   // jamais trouv� de droit explicite
        psDroit:= cstDroitModif;  // pas de droit <=> tous les droits

  finally
    reqLieu.Free;
  end;
  SortieProc;
end;


{ -------------------------------------------------------------------------------------- }
{ Retourne la clause where d'une requ�te SQL s�lectionnant toutes les sessions p�rim�es }
function clsReqGenerique.CritereSessionPerimee(
 peTopHorloge: TDateTime;   // top horloge en jours
 peDelaiVeille: integer;    // d�lai de veille en secondes (issu du param�trage)
 peDureeMaxSession: integer; // dur�e maxi d'une session en secondes (issu du param�trage)
 peDureeMaxRequete: integer) // dur�e maxi d'une requ�te en secondes (issu du param�trage)
 : string;
{ -------------------------------------------------------------------------------------- }

var
  topHorlVeille,topHorlSession,topHorlRequete: TDateTime; // TDateTime <=> Double

begin
  EntreeProc(ClassName+'.CritereSessionPerimee');
{ Calcul des temps-limites en jours (=86400 s) (c'est l'unit� utilis�e par GetDate et GetTime) }
  topHorlVeille:= peTopHorloge - peDelaiVeille/86400;  // temps avant lequel une session est en veille
  topHorlSession:= peTopHorloge - peDureeMaxSession/86400;  // temps avant lequel une session est p�rim�e
  topHorlRequete:= peTopHorloge - peDureeMaxRequete/86400;  // temps avant lequel une requ�te doit �tre consid�r�e comme arr�t�e

  result:= format ('(ReqEnCours <> 1 or TopRequete < %5.6f) '
   // On exclut les sessions ayant une requ�te en cours ayant commenc� apr�s topHorlRequete
   // (m�me si une d�connexion est demand�e)
   +'and (TopRequete < %5.6f or UtilUnique <> 0 and TopRequete < %5.6f '
   // On prend les sessions ayant fait leur derni�re requ�te avant topHorlSession
   // ou topHorlVeille pour celles en mode utilisateur unique
   +'or DcnxPrevue is not null and DcnxPrevue < %5.6f)',
   // On prend les sessions ayant d�pass� leur d�lai de d�connexion
   // NB: sans le test not null on ne pourrait pas utiliser NOT (CritereSessionPerimee) dans une requ�te
   // car son �valuation SQL renverrait NULL et pas FAUX quand DcnxPrevue est � NULL
   [topHorlRequete,topHorlSession,topHorlVeille,peTopHorloge],FormatsADN);
  SortieProc;
end;

{ -------------------------------------------------------------------------------------- }
{ Retourne la clause where d'une requ�te SQL s�lectionnant toutes les sessions ACTIVES
  et pas en veille ni p�rim�es }
function clsReqGenerique.CritereSessionActive(
 peTopHorloge: TDateTime;   // top horloge en jours
 peDelaiVeille: integer;    // d�lai de veille en secondes (issu du param�trage)
 peDureeMaxRequete: integer) // dur�e maxi d'une requ�te en secondes (issu du param�trage)
 : string;
{ -------------------------------------------------------------------------------------- }

var
  topHorlVeille,topHorlRequete: TDateTime; // TDateTime <=> Double

begin
  EntreeProc(ClassName+'.CritereSessionActive');
{ Calcul des temps-limites en jours (=86400 s) (c'est l'unit� utilis�e par GetDate et GetTime) }
  topHorlVeille:= peTopHorloge - peDelaiVeille/86400;  // temps avant lequel une session est en veille
  topHorlRequete:= peTopHorloge - peDureeMaxRequete/86400;  // temps avant lequel une requ�te doit �tre consid�r�e comme arr�t�e

  result:= format ('ReqEnCours = 1 and TopRequete >= %5.6f '
   // On inclut les sessions ayant une requ�te en cours ayant commenc� apr�s topHorlRequete
   // (m�me si elles ont d�pass� leur d�lai de d�connexion)
   +'or TopRequete >= %5.6f '
   // On prend les sessions ayant fait leur derni�re requ�te apr�s topHorlVeille ...
   +'and (DcnxPrevue is null or DcnxPrevue >= %5.6f)',
   // ... et n'ayant pas d�pass� leur d�lai de d�connexion
   [topHorlRequete,topHorlVeille,peTopHorloge],FormatsADN);
  SortieProc;
end;

{ -------------------------------------------------------------------------------------- }
{ Retourne la clause where d'une requ�te SQL s�lectionnant toutes les sessions EN VEILLE
  et pas en veille ni p�rim�es }
function clsReqGenerique.CritereSessionEnVeille(
 peTopHorloge: TDateTime;   // top horloge en jours
 peDelaiVeille: integer;    // d�lai de veille en secondes (issu du param�trage)
 peDureeMaxSession: integer; // dur�e maxi d'une session en secondes (issu du param�trage)
 peDureeMaxRequete: integer) // dur�e maxi d'une requ�te en secondes (issu du param�trage)
 : string;
{ -------------------------------------------------------------------------------------- }

var
  topHorlVeille,topHorlSession,topHorlRequete: TDateTime; // TDateTime <=> Double

begin
  EntreeProc(ClassName+'.CritereSessionEnVeille');
{ Calcul des temps-limites en jours (=86400 s) (c'est l'unit� utilis�e par GetDate et GetTime) }
  topHorlVeille:= peTopHorloge - peDelaiVeille/86400;  // temps avant lequel une session est en veille
  topHorlSession:= peTopHorloge - peDureeMaxSession/86400;  // temps avant lequel une session est p�rim�e
  topHorlRequete:= peTopHorloge - peDureeMaxRequete/86400;  // temps avant lequel une requ�te doit �tre consid�r�e comme arr�t�e

  result:= format ('(ReqEnCours <> 1 or TopRequete < %5.6f) '
   // On exclut les sessions ayant une requ�te en cours ayant commenc� apr�s topHorlRequete
   // (m�me si elles ont d�pass� leur d�lai de d�connexion)
   +'and TopRequete < %5.6f and TopRequete >= %5.6f '
   // derni�re requ�te entre topHorlVeille et TopHorlSession
   +'and UtilUnique = 0 '
   // Une session avec utilisateur unique ne peut pas �tre �tre en veille
   +'and (DcnxPrevue is null or DcnxPrevue >= %5.6f)',
   // On prend les sessions n'ayant pas d�pass� leur d�lai de d�connexion
   [topHorlRequete,topHorlVeille,topHorlSession,peTopHorloge],FormatsADN);
  SortieProc;
end;


{ -------------------------------------------------------------------------------------- }
{ V�rifie valeur de propri�t� selon le format (peFormat) }
function clsReqGenerique.VpFormatOK
 (const peValeur,peFormat: string;
 var psValeurFormatee: string;   // sortie: valeur reformat�e, stockable en base
 var psMsg : string)
 : boolean;  // renvoie TRUE si OK sinon FALSE + psMsg
{  Si format et valeur vide renvoie OK
  Cas limites du format d�cimal
  - Absence virgule => les chiffres trouv�s portent sur la partie enti�re
  - Absence chiffres partie enti�re ET d�cimale => Erreur format
{ Proc�dure remani�e par MG le 30/08/07 }
{ -------------------------------------------------------------------------------------- }

var
 valeur, calibre, chNbCaracteres, chNbCarPartieEntiere, chNbCarPartieDecimale,entiersD, decimalesD: string;
 lcal, lval, pSepDecimal, intValeur, nbEntiers, nbDecimales, iEntiers  : integer;
 reel : double; ladate: TdateTime;
 videInterdit : boolean;

 procedure addMsg(const peStr: string);
 begin
   if psMsg ='' then psMsg := peStr else psMsg := psMsg + ', ' + peStr;
 end;

begin
  EntreeProc(ClassName+'.vpFormatOK');
  psMsg  :='';
  result :=TRUE;
  valeur := trim(peValeur);
  calibre:= trim(peFormat); // format est une fonction delphi
  lcal   := length(calibre);
  if lcal=0  then
    EXCEPTIONINTERNE(defErr18);   // Format de propri�t� incorrect
  lval   := length(valeur);

(* 3.6.6  videInterdit:= (calibre[length(calibre)] = '!');          // champ obligatoirement non vide ?
  if videInterdit then
    entiers :=  midStr(calibre,2,lcal-2)
  else
    entiers:=midStr(calibre,2,lcal-1);    // extraction de la zone num�rique
*)
  { v3.6.6: le '!' n'est pas forc�ment � la fin de calibre : ex: C10!$6U }
  videInterdit:= pos('!',calibre) > 0;
  chNbCaracteres:= '';
  for iEntiers:= 2 to length(calibre) do
  begin
    if (calibre[iEntiers] in ['0'..'9']) or (calibre[iEntiers] = formatsADN.DecimalSeparator) then
      chNbCaracteres:= chNbCaracteres + calibre[iEntiers]
    else
      BREAK;
  end;

  if (lval=0) and videInterdit then
  begin
    result:= FALSE;
    psMsg:= lbErrValeurVideRefusee   // ce message contient un %s et l'appelant doit le transformer par format(...,[nom de la prop.])
  end
  else
  begin
    if (lval<>0) then      // valeur non vide : il faut la contr�ler
    begin
      case  calibre[1] of      // V�rification  du format lui-m�me
        'E','M','C','L':                       // format entier, majuscule et caract�re
          if not TryStrToInt(chNbCaracteres,nbEntiers) then
            EXCEPTIONINTERNE(defErr18);   // Format de propri�t� incorrect

        'D':                               // format D�cimal
          begin
            pSepDecimal:= pos(formatsADN.DecimalSeparator,chNbCaracteres);
            if (pSepDecimal>0) then begin
              chNbCarPartieEntiere:= midStr(chNbCaracteres,1,pSepDecimal-1);
              chNbCarPartieDecimale:= midStr(chNbCaracteres,pSepDecimal+1,length(chNbCaracteres)-pSepDecimal);
            end
            else chNbCarPartieEntiere:= chNbCaracteres; // pas de s�parateur on d�cide que c'est entier

            if chNbCarPartieEntiere='' then
              chNbCarPartieEntiere:='0';
            if chNbCarPartieDecimale='' then
              chNbCarPartieDecimale:='0';

            if not TryStrToInt(chNbCarPartieEntiere,nbEntiers) or not TryStrToInt(chNbCarPartieDecimale,nbDecimales)then
              EXCEPTIONINTERNE(defErr18);   // Format de propri�t� incorrect
            if (nbEntiers=0) and (nbDecimales=0) then
              EXCEPTIONINTERNE(defErr18)   // Format de propri�t� incorrect
            // sinon le format incoh�rent 'D.!' est valide !
          end; // cas 'D'

        'J':   // MG : contr�le supprim�: on pourra peut-�tre mettre 'J' au lieu de l'inutile JJMMAAAA

        else
          EXCEPTIONINTERNE(defErr18)   // Format de propri�t� incorrect
      end; // case v�rification du format

      { v�rification de la compatibilit� de la valeur avec le format }

      case  calibre[1] of          // Verification valeur enti�re
      'E':
        if not TryStrToInt(valeur,intValeur) then
          addMsg(lbErrValeurEntiere)
        else
          if lval>nbEntiers then
            addMsg(format(lbErrTaille,[chNbCaracteres]))
          else
            valeur:= IntToStr(intValeur);

      'D':
        begin                         // Verification valeur d�cimale
          if (pos(',',valeur) >0) then    // substitution virgule par le separateur
            valeur := StringReplace(valeur,',',formatsADN.DecimalSeparator,[]);
          if not TryStrToFloat(valeur,reel,FormatsADN) then
            addMsg(lbErrValeurDecimale)
          else
          begin                  // d�codage partie enti�re et d�cimale de la valeur
             valeur:= FloatToStr(reel,FormatsADN);
             pSepDecimal := pos(formatsADN.DecimalSeparator,valeur);
             if (pSepDecimal>0) then
             begin
               entiersD:= midStr(valeur,1,pSepDecimal-1);
               decimalesD:= midStr(valeur,pSepDecimal+1,length(valeur)-pSepDecimal);
             end
             else
             begin          // pas de s�parateur on d�cide que c'est entier
                entiersD := valeur;
                decimalesD := '';
             end;
             if (length(entiersD)> nbEntiers) then
                addMsg(format(lbErrTailleEntier,[chNbCarPartieEntiere]));
             if (length(decimalesD)> nbDecimales) then
                addMsg(format(lbErrTailledecimale,[chNbCarPartieDecimale]));
          end;
        end;

      'M':                               // Verification caract�res majuscules
        if lval>nbEntiers then
          addMsg(format(lbErrTaille,[chNbCaracteres]))
        else
          valeur:= ansiUpperCase(valeur);  // Modif MG: on accepte la saisie en minuscules !

      'C','L':                               // Verification caract�res
        if lval>nbEntiers then
          addMsg(format(lbErrTaille,[chNbCaracteres]));

      'J':                               // Verification date
        if tryStrToDate(valeur,ladate,FormatsADN) then
          // NB: TryStrToDate accepte la non-saisie de l'ann�e dans la date
          // et ajoute dans ce cas l'ann�e en cours
          valeur:= IntToStr(trunc(ladate))
        else
          addMsg(lbErrValeurDate);
      end;  // case calibre[1] of
    end;

    if psMsg = '' then
      psValeurFormatee:= valeur
    else
    begin
      psMsg:=format(lbErrValeurPropriete,[RendNonXml(peValeur),'%s',psMsg]);
      // On garde l'autre %s car l'appelant doit le transformer par format(...,[nom de la prop.])
      result:=FALSE;
    end;
  end;

  SortieProc;
end; // VpFormatOK

{ ------------------------------------------------------------------------------------------}
procedure clsReqGenerique.LitParamChaine(
 peTNumPar: array of integer;         // tableau des num�ros de param�tres � rechercher
 peTValeurDefaut: array of string;  // tabDynChaine;
 var psValParam: tabAttrib);   // tableau des valeurs lues
{ Lecture des param�tres g�n�raux dont les noms sont dans peTNomsPar }
{ ------------------------------------------------------------------------------------------}

var
  reqParGene: TADOQuery;
  indPar: integer;
  valeurLue: string;

begin
  EntreeProc(ClassName+'.LitParamChaine');
  SetLength(psValParam,length(peTNumPar));
  reqParGene:= CreeADOQuery();
  try
//    reqParGene.Connection:= ConnexionBD;
    with reqParGene do
    begin
      SQL.Add(format('select * from Parametre where Numero = %d ',[peTNumPar[0]]));
      for indPar:= 1 to high(peTNumPar) do
        SQL.Add(format('or Numero = %d ',[peTNumPar[indPar]]));
      Open;
      for indPar:= 0 to high(peTNumPar) do
      begin
        if locate('Numero',peTNumPar[indPar],[loCaseInsensitive]) then
          valeurLue:= FieldByName('Valeur').asString
        else
          valeurLue:= '';
        if valeurLue = '' then
          psValParam[indPar]:= peTValeurDefaut[indPar]
        else
          psValParam[indPar]:= valeurLue;
      end;
    end;
  finally
    reqParGene.Free;
  end;
  SortieProc;
end;


{ ------------------------------------------------------------------------------------------}
procedure clsReqGenerique.LitParamEntiers(
 peTNumPar: array of integer;         // tableau des num�ros de param�tres � rechercher
 peTValeurDefaut: array of integer;
 var psValParam: tabDynEntier);   // tableau des valeurs lues
{ Lecture des param�tres g�n�raux dont les num�ros sont dans peTNumPar }
{ ------------------------------------------------------------------------------------------}

var
  reqParGene: TADOQuery;
  indPar: integer;
  chValeurLue: string;

begin
  EntreeProc(ClassName+'.LitParamEntiers');
  reqParGene:= CreeADOQuery();
  SetLength(psValParam,length(peTNumPar));
  try
//    reqParGene.Connection:= ConnexionBD;
    with reqParGene do
    begin
      SQL.Add(format('select Numero,Valeur from Parametre where Numero = %d ',[peTNumPar[0]]));
      for indPar:= 1 to high(peTNumPar) do
        SQL.Add(format('or Numero = %d',[peTNumPar[indPar]]));
      Open;
      for indPar:= 0 to high(peTNumPar) do
      begin
        { Modifi� v3.5.0 pour accepter des valeurs � NULL ou vides }
        if locate('Numero',peTNumPar[indPar],[loCaseInsensitive]) then
          chValeurLue:= FieldByName('Valeur').AsString
        else
          chValeurLue:= '';
        psValParam[indPar]:= strToIntDef(chValeurLue,peTValeurDefaut[indPar]);
      end;
    end;
  finally
    reqParGene.Free;
  end;
  SortieProc;
end;

{ ------------------------------------------------------------------------------------------}
{ Ajout d'une action symbolique dans l'historique globale � un site }
function clsReqGenerique.AjouteActionGlobale(
 peOperation: TOperationR3Web)
 : integer;
{ ------------------------------------------------------------------------------------------}

var
  rqInsertion: TADOQuery;
  texteSQL: TStringList;

begin
  EntreeProc(ClassName+'.AjouteActionGlobale');
  RqInsertion:= CreeADOQuery();
  texteSQL:= TStringList.Create;

  try
    with rqInsertion do
    begin
      SQL.Add('insert into Action (Id_uti,Id_ses,Etat,Operation,Moment,Options,IdSite)');
      SQL.Add(format('values (%d,%d,''%s'',%d,%5.6f,%d,%s)',
       [IdUtilisateur,NumSession,InfoEtatAction[eExecutee].Code,ord(peOperation),Date+GetTime,Options,IdSite],formatsADN));
      SQLRecupereIdent(texteSQL,'identAction',true);
      SQL.AddStrings(texteSQL);    // Ajoute le code SQL cr�� par SQLRecupereIdent
      SQL.Add ('select @identAction [Resultat]');
      Open;
      result:= FieldByName('Resultat').AsInteger;
    end;

  finally
    texteSQL.Free;
    rqInsertion.Free;
  end;
  SortieProc;
end;


{ ------------------------------------------------------------------------------------------}
{ Ajout d'une action dans l'historique - utilis�e pour les op�rations de cablage et placement }
function clsReqGenerique.AjouteActionCablage(
 peCodeOper: TOperationR3Web;   // code op�ration
 peCategorie: TCategorie;     // code cat�gorie d'objet
 peIdObjet: integer;       // identifiant d'objet
 peNomObjet: string;        // nom de l'objet
{peActionContexte: integer = 0;    // supprim� en v3.5.3a - tient compte de NumActionContexte � la place }
 peComm: string = '';      // Commentaire �ventuel
 peNbCnx: integer = 0;   // Nombre de connexions pour les op�rations autres que maj de composant
 peNomCncDep: string = '';   // nom du premier connecteur de d�part
 peIdCncDep: integer = 0;   // identifiant du premier connecteur de d�part
 peCnxDep: string = '';   // premi�re connexion de d�part ou bien premier fil de fonction retir� (si peNomCncDep = '')
 peNomDerCncDep: string = '';   // nom du dernier connecteur de d�part
 peDerCnxDep: string = '';   // premi�re connexion de d�part
 peNomCncArr: string = '';   // nom du premier connecteur d'arriv�e
 peIdCncArr: integer = 0;   // identifiant du premier connecteur d'arriv�e
 peCnxArr: string = '';   // premi�re connexion d'arriv�e
 peLigneeDepart: TStringList = nil;   // LT [+ GE] + eqt de d�part
 peCategLigneeDepMax: TCategorie = eEquipement;    // Cat�gorie du dernier �l�ment de la lign�e (pour d�c�blage sur tout un GE ou tout un LT)
 peIdEqtDep: integer = 0;   // identifiant de l'�quipement de d�part
 peLigneeArrivee: TStringList = nil;      // LT [+ GE] + eqt d'arriv�e si l'op�ration en poss�de
 peIdEqtArr: integer = 0;   // identifiant de l'�quipement d'arriv�e
 peFilCabFonc: integer = 0;    // premier fil de fonction ou de c�ble
 peLongueur: integer = -1;    // longueur (de brassage)
 peValeursProprietes: TIdentValeurFormat = nil)
 : integer;        // renvoie le num�ro d'action g�n�r�e
{ ------------------------------------------------------------------------------------------}

var
  topHorloge: TDateTime;
  nomLtDep,nomGeDep,NomEqtDep,nomLtArr,nomGeArr,nomEqtArr: string;
  texteSQL: TStringList;
  rqInsertion: TADOQuery;

begin
  EntreeProc(ClassName+'.AjouteActionCablage');

  (* SUPPRIME v3.5.3: on utilise NumActionContexte
   { Si on a fourni une valeur de peActionContexte � 0 mais qu'on est dans une sous-requ�te
    (induite par une autre requ�te), alors on prend comme contexte le n� d'action de la requ�te m�re }
  if (peActionContexte = 0) and (RequeteAppelante <> nil) then
    peActionContexte:= RequeteAppelante.NumAction;
*)

  if EnregHistorique then
  begin
    texteSQL:= TStringList.Create;
    rqInsertion:= CreeADOQuery();      // v3.5.0: remplace Query car il est dangereux d'y mettre trop de choses

    try
      topHorloge:= Date+GetTime;
      with rqInsertion do
      begin
        SQL.Clear;
        SQL.Add('insert into Action (Id_uti,Id_ses,Id_Odt,Etat,Operation,Categorie,IdObjet,NomObjet,');
        SQL.Add('ActionContexte,Moment,Comm,NbCnx,PremierFil,IdSite,Options,IdActIni,Longueur)');
        SQL.Add(format('values (%d,%d,',[IdUtilisateur,NumSession]));
        { Maj de Id_Odt,Etat }
        if OdtActif > 0 then
          SQL.Add(format('%d,''%s'',',[OdtActif,InfoEtatAction[ePrevue].Code]))   // Mode diff�r�
        else
          if OdtExecute > 0 then
            SQL.Add(format('%d,''%s'',',[OdtExecute,InfoEtatAction[eExecutee].Code])) // Mode ex�cution ODT
          else
            SQL.Add(format('NULL,''%s'',',[InfoEtatAction[eExecutee].Code]));  // Mode direct
            // (laisser le champ � NULL pour ne pas contrarier la FK)
        SQL.Add(format('%d,',[ord(peCodeOper)]));  // Operation
        if peIdObjet = 0 then  // exemple: retrait de fonctions � partir d'un point
        begin
          if peCodeOper = eRetirerFonction then
            SQL.Add(format('''%s'',',[InfoCateg[eFonction].Code]))  // Categorie
          else
            SQL.Add('NULL,');  // Cat�gorie � NULL (suppression de brassage)
          SQL.Add('NULL,NULL,')  // IdObjet,NomObjet
        end
        else
          SQL.Add(format('''%s'',%d,%s,',
           [InfoCateg[peCategorie].Code,peIdObjet,quotedStr(peNomObjet)]));  // IdObjet,NomObjet
        if NumActionContexte = 0 then  // laisser le champ � NULL pour ne pas contrarier la FK
          SQL.Add(format('NULL,%5.6f,%s,',
           [topHorloge,quotedStr(peComm)],FormatsADN))
        else
          SQL.Add(format('%d,%5.6f,%s,',
           [NumActionContexte,topHorloge,quotedStr(peComm)],FormatsADN));
        if peNbCnx = 0 then  // NbCnx n'a pas de sens ou est inconnu (ex: retrait sur tout un eqt)
          SQL.Add('NULL,')
        else
          SQL.Add(format('%d,',[peNbCnx]));
        if peFilCabFonc > 0 then
          SQL.Add(format('%d,',[peFilCabFonc]))
        else
          SQL.Add('NULL,');
        if IdSite = '' then
          SQL.Add(format('NULL,%d,',[Options]))
        else
          SQL.Add(format('%s,%d,',[IdSite,Options]));

        if (OdtExecute > 0) and (NumActionContexte = 0) then   // Mode ex�cution ODT (en mode simul� on n'appelle jamais cette proc)
          SQL.Add(format('%d,',[NumAction]))   // IdActIni
        else
          SQL.Add('NULL,');

        if peLongueur = - 1 then      // ajout v3.4.6 (802) (et ")" enlev�e aux 2 SQL.Add pr�c�dents)
          SQL.Add('NULL)')
        else
          SQL.Add(format('%d)',[peLongueur]));

        SQLRecupereIdent(texteSQL,'identAction',true);
        SQL.AddStrings(texteSQL);    // Ajoute le code SQL cr�� par SQLRecupereIdent

        if (peLigneeDepart = nil) or (peLigneeDepart.Count = 0) then
          if peIdCncDep <> 0 then
            AncetresConnecteur(peIdCncDep,nomLtDep,nomGedep,nomEqtDep)
          else
          begin
            nomLtDep:= '';
            nomGeDep:= '';
            nomEqtDep:= '';
          end
        else
        begin
          case peCategLigneeDepMax of
          eTerminaison:
            begin
              nomLtDep:= '';
              nomGedep:= '';
              nomEqtDep:= peLigneeDepart.Strings[0]   // terminaison <=> lign�e avec juste un �quipement
            end;
          eLocalTechnique:
            begin
              nomLtDep:= peLigneeDepart.Strings[0];   // la lign�e commence par le LT
              nomGeDep:= '';
              nomEqtDep:= '';
            end;
          eGroupe:
            begin
              nomLtDep:= peLigneeDepart.Strings[0];   // la lign�e commence par le LT
              nomGeDep:= peLigneeDepart.Strings[1]
            end;
          eEquipement:
            begin
              nomLtDep:= peLigneeDepart.Strings[0];   // la lign�e commence par le LT
              if peLigneeDepart.Count = 2 then
              begin
                nomGeDep:= '';
                nomEqtDep:= peLigneeDepart.Strings[1]
              end
              else  // peLigneeDepart.Count = 3
              begin
                nomGeDep:= peLigneeDepart.Strings[1];
                nomEqtDep:= peLigneeDepart.Strings[2];
              end
            end;
          end;     // case peCategLigneeDepMax
        end;

        if (peLigneeArrivee = nil) or (peLigneeArrivee.Count = 0) then
          if peIdCncArr <> 0 then
            AncetresConnecteur(peIdCncArr,nomLtArr,nomGeArr,nomEqtArr)
          else
          begin
            nomLtArr:= '';
            nomGeArr:= '';
            nomEqtArr:= '';
          end
        else
        begin    // La lign�e d'arriv�e, si elle est fournie, va toujours jusqu'� l'�quipement ou la terminaison
          { Rappel: si terminaison, c'est le seul �lt de la lign�e }
          nomEqtArr:= peLigneeArrivee.Strings[pred(peLigneeArrivee.Count)];
          case peLigneeArrivee.Count of
            3:
            begin
              nomLtArr:= peLigneeArrivee.Strings[0];
              nomGeArr:= peLigneeArrivee.Strings[1];
            end;
            2:
            begin
              nomLtArr:= peLigneeArrivee.Strings[0];
              nomGeArr:= '';
            end;
            1:
            begin  // seul l'�quipement est pr�cis� (c'est une terminaison)
              nomGeArr:= '';
              nomLtArr:= '';
            end;
          end;
        end;

        if (nomLtDep <> '') or (nomEqtDep <> '') then   // op�ration avec pr�cision d'un d�part
        begin
          SQL.Add('insert into DetailAction (');
          SQL.Add('Id_act,NomLT,NomGE,NomEqt,NomCnc,Id_cnc,Cnx,NomDerCnc,DerCnx,Origine');
          if peIdEqtDep <> 0 then
            SQL.Add(',Id_eqt');
          SQL.Add(')');
          SQL.Add('values (@identAction');

          if nomLtDep = '' then
            SQL.Add(',NULL')
          else
            SQL.Add(format(',%s',[quotedStr(nomLtDep)]));

          if nomGeDep = '' then
            SQL.Add(',NULL')
          else
            SQL.Add(format(',%s',[quotedStr(nomGeDep)]));

          if nomEqtDep = '' then
            SQL.Add(',NULL')
          else
            SQL.Add(format(',%s',[quotedStr(nomEqtDep)]));

          if peNomCncDep = '' then
            SQL.Add(',NULL')
          else
            SQL.Add(format(',%s',[quotedStr(peNomCncDep)]));

          if peIdCncDep = 0 then
            SQL.Add(',NULL,NULL')     // v3.4.6a (804) Si peIdCncDep = 0 on ignore peCnxDep
          else
          begin
            SQL.Add(format(',%d',[peIdCncDep]));
            if peCnxDep = '' then
              SQL.Add(',NULL')
            else
              SQL.Add(format(',%s',[quotedStr(peCnxDep)]));
          end;
          if peNomDerCncDep = '' then
            SQL.Add(',NULL,NULL')       // v3.4.6a (804) Si peIdCncDep = 0 on ignore peCnxDep (qui peut avoir �t� maj!)
          else
          begin
            SQL.Add(format(',%s',[quotedStr(peNomDerCncDep)]));
            if peDerCnxDep = '' then
              SQL.Add(',NULL')
            else
              SQL.Add(format(',%s',[quotedStr(peDerCnxDep)]));
          end;
          SQL.Add(',1');    // Indicateur "c�t� d�part"

          if peIdEqtDep <> 0 then
            SQL.Add(format(',%d',[peIdEqtDep]));
          SQL.Add(')');

// MG 3.4.1       if peLigneeArrivee <> nil then
          if nomEqtArr <> '' then    // MG 3.4.1
          begin    // ce n'est pas le c�blage d'un demi-fil de c�ble ou un d�c�blage ou retrait � partir d'un plage de d�part
            SQL.Add('insert into DetailAction (');
            SQL.Add('Id_act,NomLT,NomGE,NomEqt,NomCnc,Id_cnc,Cnx,Origine');
            if peIdEqtArr <> 0 then
              SQL.Add(',Id_eqt');
            SQL.Add(')');

            SQL.Add('values (@identAction');

            if nomLtArr = '' then
              SQL.Add(',NULL')
            else
              SQL.Add(format(',%s',[quotedStr(nomLtArr)]));

            if nomGeArr = '' then
              SQL.Add(',NULL')
            else
              SQL.Add(format(',%s',[quotedStr(nomGeArr)]));

            if nomEqtArr = '' then
              SQL.Add(',NULL')
            else
              SQL.Add(format(',%s',[quotedStr(nomEqtArr)]));

            if peNomCncArr = '' then
              SQL.Add(',NULL')
            else
              SQL.Add(format(',%s',[quotedStr(peNomCncArr)]));

            if peIdCncArr = 0 then
              SQL.Add(',NULL,NULL')
            else
            begin
              SQL.Add(format(',%d',[peIdCncArr]));
              if peCnxArr = '' then
                SQL.Add(',NULL')
              else
                SQL.Add(format(',%s',[quotedStr(peCnxArr)]));
            end;

            SQL.Add(',0');    // Indicateur "c�t� arriv�e"

            if peIdEqtArr <> 0 then
              SQL.Add(format(',%d',[peIdEqtArr]));
            SQL.Add(')');
          end;
        end;

        if (OdtActif > 0) then
          { On fait passer le champ IdActIni � la valeur de Id_act :
          utilis� pour ordonner la fiche de travaux }
          SQL.Add('update Action set IdActIni = @identAction where Id_act = @identAction');

        SQL.Add ('select @identAction as resultat');
        Open;      // et pas ExecSQL qui ne permet pas de r�cup�rer un r�sultat
        result:= FieldByName('resultat').asInteger;   // retourne le n� d'action g�n�r�e
      end;    // with rqInsertion
    finally
      rqInsertion.Free;
      texteSQL.Free;
    end;
  end
  else
    result:= 0;   // pour �viter avertissement compilateur
  SortieProc;
end;

{ ------------------------------------------------------------------------------------------}
{ Renvoie le SQL d'ajout d'action historique, destin� � �tre ex�cut� plus tard
 (group� avec d'autres instructions SQL }

procedure clsReqGenerique.SQLAjouteActionCompo(
 var peTexteSQL: TStringList;  // Commandes SQL � mettre � jour
 peCodeOper: TOperationR3Web;   // code op�ration
 peCategorie: TCategorie;     // code cat�gorie d'objet
 peIdObjet: integer;       // identifiant d'objet
 peNomObjet: string;     // nom de l'objet
{peActionContexte: integer = 0;    // supprim� en v3.5.3a - tient compte de NumActionContexte � la place }
 peComm: string = '';      // Commentaire �ventuel
 peNomLt: string = '';    // nom du LT d'appartenance si l'objet cr�� est un �quipement
 peNomGe: string = '');   // nom du GE d'appartenance si l'objet cr�� est un eqt dans un GE
{ ------------------------------------------------------------------------------------------}

var
  topHorloge: TDateTime;

begin
  EntreeProc(ClassName+'.SQLAjouteActionCompo');
  if EnregHistorique then
  begin
    topHorloge:= Date+GetTime;
(* Supprim� v3.5.3 - on prend NumActionContexte � la place
    if (peActionContexte = 0) and (RequeteAppelante <> nil) then
      peActionContexte:= RequeteAppelante.NumAction;
*)
    with peTexteSQL do
    begin
      Add('insert into Action');
      Add('(Id_uti,Id_ses,Id_Odt,Etat,Operation,IdSite,Categorie,IdObjet,NomObjet,');
      Add('ActionContexte,Options,Moment,Comm');

      if OdtExecute > 0 then   // Mode ex�cution ODT (en mode simul� on n'appelle jamais cette proc)
        Add(',IdActIni');

      case peCategorie of
        eEquipement:
          Add(',NomLT,NomGE)');
        eGroupe:
          Add(',NomLT)')
        else
          Add(')');
      end;

      Add(format('values (%d,%d',[IdUtilisateur,NumSession]));

      if OdtActif > 0 then
        // NB: OdtExecute est alors = 0 sinon on ne doit pas appeler cette proc. (mode simul�)
        Add(format(',%d,''%s''',[OdtActif,InfoEtatAction[ePrevue].Code]))
      else
        if OdtExecute > 0 then
          Add(format(',%d,''%s''',[OdtExecute,InfoEtatAction[eExecutee].Code]))
        else
          // laisser le champ � NULL pour ne pas contrarier la FK
          Add(format(',NULL,''%s''',[InfoEtatAction[eExecutee].Code]));    // Id_Odt,Etat

      if IdSite = '' then
        Add(format(',%d,NULL',[ord(peCodeOper)]))  // Op�ration sans site
      else
        Add(format(',%d,%s',[ord(peCodeOper),IdSite]));  // NB: ne marche pas si IdSite = ''
      if peIdObjet = 0 then     // Action sans objet (?)
        Add(',NULL,NULL,NULL')
      else
        Add(format(',''%s'',%d,%s',   // categorie,IdObjet,nomObjet
         [InfoCateg[peCategorie].Code,peIdObjet,quotedStr(peNomObjet)]));
      if NumActionContexte = 0 then  // laisser le champ � NULL pour ne pas contrarier la FK
        Add(format(',NULL,%d,%5.6f,%s',
         [Options,topHorloge,quotedStr(peComm)],FormatsADN))
      else
        Add(format(',%d,%d,%5.6f,%s',[NumActionContexte,Options,topHorloge,quotedStr(peComm)],
         FormatsADN));

      if OdtExecute > 0 then   // Mode ex�cution ODT (en mode simul� on n'appelle jamais cette proc)
        Add(format(',%d',[NumAction]));   // IdActIni

      case peCategorie of
        eEquipement:
          Add(format(',%s,%s)',[quotedStr(peNomLt),quotedStr(peNomGe)]));
        eGroupe:
          Add(format(',%s)',[quotedStr(peNomLt)]))
        else
          Add(')');
      end;

      SQLRecupereIdent(peTexteSQL,'identAction');

      if OdtActif > 0 then
        { On fait passer le champ IdActIni � la valeur de Id_act :
          utilis� pour ordonner la fiche de travaux }
        Add('update Action set IdActIni = @identAction where Id_act = @identAction');

    end;    // with peTexteSQL
  end;
  SortieProc;   // d�plac� en 3.6.0 (plus correct!)
end;


{ ------------------------------------------------------------------------------------------}
{ G�n�re dans peTexteSQL les commandes d'ajout en historique d'une cr�ation d'objet }
{ La diff�rence avec les autres actions est qu'on ne conna�t pas encore l'identifiant
  de l'objet et donc on fait passer un nom de variable SQL qui sera � jour au moment de
  l'ex�cution de la commande SQL }

procedure clsReqGenerique.SQLAjouteCreationCompo(
 var peTexteSQL: TStringList;  // Commandes SQL � mettre � jour
 peCategorie: TCategorie;     // code cat�gorie d'objet
 peNomVarSQL: string;    // nom de la variable SQL (comman�ant par @) contenant l'identifiant d'objet
                         // ou bien valeur de l'ident elle-m�me sous forme de cha�ne
 peNomObjet: string;     // nom de l'objet cr��
{peActionContexte: integer = 0;    // supprim� en v3.5.3a - tient compte de NumActionContexte � la place }
 peComm: string = '';      // Commentaire �ventuel
 peNomLt: string = '';    // nom du LT d'appartenance si l'objet cr�� est un �quipement
 peNomGe: string = '');   // nom du GE d'appartenance si l'objet cr�� est un eqt dans un GE

 { ------------------------------------------------------------------------------------------}

var
  topHorloge: TDateTime;

begin
  EntreeProc(ClassName+'.SQLAjouteCreationCompo');
  if EnregHistorique then
  begin
    topHorloge:= Date+GetTime;
(* Supprim� v3.5.3 - on pend NumActionContexte � la place
    if (peActionContexte = 0) and (RequeteAppelante <> nil) then
      peActionContexte:= RequeteAppelante.NumAction;
*)
    with peTexteSQL do
    begin
      Add('insert into Action');
      Add('(Id_uti,Id_ses,Id_Odt,Etat,Operation,IdSite,Categorie,IdObjet,NomObjet,ActionContexte,Moment,Comm');

      if OdtExecute > 0 then   // Mode ex�cution ODT (en mode simul� on n'appelle jamais cette proc)
        Add(',IdActIni');

      case peCategorie of
        eEquipement:
          Add(',NomLT,NomGE)');
        eGroupe:
          Add(',NomLT)')
        else
          Add(')');
      end;

      Add(format('values (%d,%d',[IdUtilisateur,NumSession]));

      if OdtActif > 0 then
        // NB: OdtExecute est alors = 0 sinon on ne doit pas appeler cette proc. (mode simul�)
        Add(format(',%d,''%s''',[OdtActif,InfoEtatAction[ePrevue].Code]))
      else
        if OdtExecute > 0 then
          Add(format(',%d,''%s''',[OdtExecute,InfoEtatAction[eExecutee].Code]))
        else
          // laisser le champ � NULL pour ne pas contrarier la FK
          Add(format(',NULL,''%s''',[InfoEtatAction[eExecutee].Code]));    // Id_Odt,Etat

      { Attention : la variable de nom peNomVarSQL doit �tre � jour et non null au moment de l'ex�cution }
       if IdSite = '' then
        Add(format(',%d,NULL,''%s'',%s,%s',   // Operation,IdSite,Categorie,IdObjet,NomObjet
         [ord(eCreerComposant),InfoCateg[peCategorie].Code,peNomVarSQL,quotedStr(peNomObjet)]))
      else
        Add(format(',%d,%s,''%s'',%s,%s',   // Operation,IdSite,Categorie,IdObjet,NomObjet
         [ord(eCreerComposant),IdSite,InfoCateg[peCategorie].Code,peNomVarSQL,quotedStr(peNomObjet)]));
      if NumActionContexte = 0 then  // laisser le champ � NULL pour ne pas contrarier la FK
        Add(format(',NULL,%5.6f,%s',
         [topHorloge,quotedStr(peComm)],FormatsADN))
      else
        Add(format(',%d,%5.6f,%s',[NumActionContexte,topHorloge,quotedStr(peComm)],
         FormatsADN));

      if OdtExecute > 0 then   // Mode ex�cution ODT (en mode simul� on n'appelle jamais cette proc)
        Add(format(',%d',[NumAction]));   // IdActIni

      case peCategorie of
        eEquipement:
          Add(format(',%s,%s)',[quotedStr(peNomLt),quotedStr(peNomGe)]));
        eGroupe:
          Add(format(',%s)',[quotedStr(peNomLt)]))
        else
          Add(')');
      end;

      { Ajoute le code correspondant � la lecture du dernier identifiant g�n�r� }
      SQLRecupereIdent(peTexteSQL,'identAction');

      if OdtActif > 0 then
        { On fait passer le champ IdActIni � la valeur de Id_act :
          utilis� pour ordonner la fiche de travaux }
        Add('update Action set IdActIni = @identAction where Id_act = @identAction');
    end;
  end;
  SortieProc;
end;


{ ----------------------------------------------------------------------------- }
{ Retrouve une valeur dans un tableau dynamique d'entiers                       }
function clsReqGenerique.TrouveEntier(
 peValCherchee: integer;   // valeur enti�re recherch�e
 peTabDyn: TabDynEntier;   // tableau dynamique d'entiers dans lequel chercher
 var psIndiceTrouve: integer)       // indice de la valeur si trouv�e
 : boolean;  // true si valeur trouv�e, false sinon
{ ----------------------------------------------------------------------------- }

var
  ind: integer;

begin
  EntreeProc(ClassName+'.TrouveDynEntier');
  result:= false;
  for ind:= 0 to high(peTabDyn) do
    if peTabDyn[ind] = peValCherchee then
    begin
      psIndiceTrouve:= ind;
      result:= true;
      BREAK
    end;
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
{ Cherche une valeur dans un tableau dynamique
et l'ajoute au tableau si elle n'est pas trouv�e                                }
function clsReqGenerique.RecenseEntier(
 peValCherchee: integer;   // valeur enti�re � recenser
 var pesTabDyn: TabDynEntier)   // tableau dynamique d'entiers dans lequel chercher
 : boolean;   // Renvoie true s'il a fallu augmenter le tableau,
               // false si la valeur y �tait d�j�

var indUtile: integer;

begin
  EntreeProc(ClassName+'.RecenseEntier');
  if TrouveEntier(peValCherchee,pesTabDyn,indUtile) then
    result:= false
  else
  begin
    setLength(pesTabDyn,succ(length(pesTabDyn)));
    pesTabDyn[high(pesTabDyn)]:= peValCherchee;
    result:= true
  end;
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
function clsReqGenerique.IdToNom(const peCateg: TCategorie;
 peIdent: integer;
 peRendreNonXml: boolean = false)   // 3.5.0c ()
 : string;
{ Selon la cat�gorie, r�cup�re un nom en fonction d'un identifiant }
{ ----------------------------------------------------------------------------- }
var ADOQuery1 : TADOQuery;

begin
  EntreeProc(ClassName+'.idToNom');
  ADOQuery1 :=  CreeADOQuery();
  with ADOQuery1 do
    try
//      Connection := ConnexionBD;
      SQL.Text:= 'select nom from '+  NomTableCompo[peCateg] + ' where ' +
      NomChampIdent(peCateg) + '=' + intTostr(peIdent) ;
      open;
      if Eof then
        result := ''
      else
        if peRendreNonXml then
          result := RendNonXml(Fields[0].asString)
        else
          result := Fields[0].asString;
    finally
      Free;
    end;
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
{ (v3.6.0) Selon la cat�gorie, trouve un composant ayant un type li� � la table Type en fonction d'un identifiant }
function clsReqGenerique.TrouveTypeComposant(
 const peCateg: TCategorie;     // cat�gorie du composant (L,R,G,E,T,B,C,F)
 const peIdent: integer)        // identifiant du composant
 : integer;        // identifiant de type du composant
{ ----------------------------------------------------------------------------- }

begin
  EntreeProc(ClassName+'.TrouveTypeComposant');
  if not (peCateg in [eLieuSimple,eLocalTechnique,eGroupe,eEquipement,eTerminaison,eFonction,eCable,eBoiteNoire]) then
    EXCEPTIONINTERNE(defErr222);
  with Query do
  begin
    SQL.Clear;
    SQL.Add(format('select Id_type from %s where %s = %d',
     [NomTableCompo[peCateg],NomChampIdent(peCateg),peIdent]));
    Open;
    if Eof then
      result:= 0
    else
      result:= FieldByName('Id_type').AsInteger;
  end;
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
function clsReqGenerique.NomComposant:
 string;  // Point d'entr�e pour la requ�te RqNomComposant (v3.4.0)
 { Modifi�e en 3.5.0 }
{ ---------------------------------------------------------------------------- }
var
  strCateg: string;
  categCompo: TCategorie;

begin
  EntreeProc(ClassName+'.NomComposant');
  strCateg:= ValeurParam(defAttCateg);
  if strCateg = '' then
    EXCEPTIONINTERNE(defErr201,'('+defAttCateg+')');
  categCompo:= DonneCategorie(strCateg[1]);
  with Query do
  begin
    SQL.Clear;
    if categCompo = eTerminaison then    // v3.5.3: il faut chercher des donn�es sp�ciales
    begin
      SQL.Add('select top 1 C.Nom NomC,S.Nom NomS,S.Id_type TypeS,S.Id_lieu NumS');
      SQL.Add(',C.Id_type TypeC,T.IdTypeCab,Cn.Nom NomCnc');
    end
    else
      SQL.Add('select C.Nom NomC,S.Nom NomS,S.Id_type TypeS,S.Id_lieu NumS');
    SQL.Add(format('from %s C',[NomTableCompo[categCompo]]));

    if (categCompo = eEquipement) or (categCompo = eTerminaison) then
    begin
      SQL.Add(format('join %s L on L.Id_lieu = C.Id_lieu',[NomTableCompo[eLieuSimple]]));
      SQL.Add(format('join %s S on S.Id_lieu = L.IdSite',[NomTableCompo[eLieuSimple]]));
      if categCompo = eTerminaison then   { v3.5.3 }
      begin
        { type de c�ble associ� � ce type de terminaison }
        SQL.Add('join Type T on T.Id_type = C.Id_type');
        { connecteur (le premier car select top 1) dans la terminaison (pour c�blage rapide) }
        SQL.Add(format('join %s Cn on Cn.Id_eqt = C.Id_eqt',[NomTableCompo[eConnecteur]]));   // v3.5.3b (1165)
      end;
    end
    else
      SQL.Add(format('join %s S on S.Id_lieu = C.IdSite',[NomTableCompo[eLieuSimple]]));
    SQL.Add(format('where C.%s = %s',[NomChampIdent(categCompo),ValeurParam(defAttNumero)]));
    Open;

    { Donn�es relatives au site }
    result:= RemplitBalise(defBalNom,FieldByName('NomS').AsString)
     + RemplitBalise(defBalNum,FieldByName('NumS').AsString);
    if FieldByName('TypeS').AsInteger = cstIdTypeLieuIntersite then
      result:= RemplitBalise(defBalSite,result,[defAttIntersite],['O'])
    else
      result:= RemplitBalise(defBalSite,result);  // pas d'attribut intersite
    result:= result + RemplitBalise(defBalNom,FieldByName('NomC').AsString);

    { v3.5.3 - Donn�es additionnelles pour terminaisons }
    if categCompo = eTerminaison then
    begin
      result:= result + RemplitBalise(defBalCnc,RendNonXml(FieldByName('NomCnc').AsString));
      result:= result + RemplitBalise(defBalNumType,FieldByName('TypeC').AsString);
      if not FieldByName('IdTypeCab').IsNull then
        result:= result + RemplitBalise(defBalTypeCableAssocie,FieldByName('IdTypeCab').AsString);
    end;

    result:= RemplitMessage(eMessInfo,result)
  end;
  SortieProc
end;

{ ----------------------------------------------------------------------------- }
function clsReqGenerique.NumeroSite(
 peNomSite: string)    // nom du site
 : integer;
{ Num�ro d'un lieu de niveau 1 (site) en fonction de son nom }
{ ----------------------------------------------------------------------------- }

begin
  EntreeProc(ClassName+'.NumeroSite');
  with Query do
  begin
    SQL.Clear;
    SQL.Add(format('select Id_lieu from %s S join Type T on T.Id_type=S.Id_type',
     [NomTableCompo[eLieuSimple]]));
    SQL.Add(format('where S.Nom = %s and T.NumNiveau = 1',[QuotedStr(peNomSite)]));
    Open;
    if Eof then  // Site non trouv� (d�phasage par renomination par exemple)
      EXCEPTIONEXTERNE(format(lbErrSiteIntrouvable,[DonneNomNiveauSite,peNomSite]),true)
    else
      result:= FieldByName('Id_lieu').AsInteger;
  end;
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
{ Renvoie l'identifiant d'un composant dont le nom est unique dans un site
  (c�ble, LT, terminaison, BN, fonction)                                        }
function clsReqGenerique.NumeroComposant(
 const peCateg: TCategorie;  // cat�gorie de composant
 const peNom: string;    // nom du composant
 peNumSite: integer;  // identifiant de site
 var psNumType: integer;  // maj l'identifiant de type du composant
 var psNumLieu: integer)  // maj le lieu d'appartenance pour les terminaisons
 : integer;  // Identifiant
{ ----------------------------------------------------------------------------- }
var ADOQuery1: TADOQuery;

begin
  EntreeProc(ClassName+'.NumeroComposant');
  ADOQuery1:= CreeADOQuery();
  with ADOQuery1 do
    try
      SQL.Add(format('select C.%s,C.Id_type',[NomChampIdent(peCateg)]));
      if peCateg = eTerminaison then
        SQL.Add(',C.Id_lieu');
      SQL.Add(format('from %s C',[NomTableCompo[peCateg]]));
      if peCateg = eTerminaison then   // cas sp�cial: la colonne IdSite est � chercher dans le lieu p�re
        SQL.Add(format('join %s L on L.Id_lieu=C.Id_lieu',[NomTableCompo[eLieuSimple]]));
      SQL.Add(format('where C.Nom = %s',[quotedStr(peNom)]));
      if peCateg = eTerminaison then
        SQL.Add(format('and L.IdSite = %d',[peNumSite]))
      else
        SQL.Add(format('and C.IdSite = %d',[peNumSite]));
      Open;
      if Eof then
        result:= 0
      else
      begin
        psNumType:= FieldByName('Id_type').AsInteger;
        if peCateg = eTerminaison then
          psNumLieu:= FieldByName('Id_lieu').AsInteger;
        result:= FieldByName(NomChampIdent(peCateg)).asInteger;
      end;
    finally
      Free;
    end;
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
{ Recherche du num�ro de l'�quipement d'un connecteur                           }
function clsreqGenerique.NumeroEquipement(
 peId_cnc: integer)     // Identifiant du connecteur � rechercher
 : integer;  // renvoie le num�ro d'�quipement auquel il appartient, 0 si pas trouv�
{ ----------------------------------------------------------------------------- }

var
  CursAncetres: TADOQuery;

begin
  EntreeProc(ClassName+'.NumeroEquipement');
  CursAncetres:= CreeADOQuery();
  with CursAncetres do
    try
//      Connection := ConnexionBD;
      SQL.Add(format('select Id_eqt from %s C where C.Id_cnc =  %d',
       [NomTableCompo[eConnecteur],peId_cnc]));
      Open;

      if Eof then
        result:= 0
      else
        result:= FieldByName('Id_eqt').AsInteger;

    finally
      Free;
    end;
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
{ Recherche d'informations textuelles sur un connecteur � partir de son identifiant }
function clsReqGenerique.LigneeCnc(
 peId_cnc: integer;  // Identifiant du connecteur � rechercher
 var psLigneeAffichee: string;  // noms du site (si besoin), du LT, du GE (s'il existe) et de l'�quipement si c'est un cnc d'�quipement
                         // nom de la terminaison s'il s'agit d'un cnc de terminaison
                         // cha�ne cod�e en "non XML"
 var psTypeCnc: string)  // nom du type de connecteur
 : string;
{ Renvoie le nom du connecteur }
{ ----------------------------------------------------------------------------- }

var
  CursAncetres: TADOQuery;
  nomLt: string;

begin
  EntreeProc(ClassName+'.LigneeCnc');
  CursAncetres:= CreeADOQuery();
  with CursAncetres do
    try
//      Connection := ConnexionBD;
      SQL.Add('select C.Nom as NomC, E.Nom as NomE, PE.Nom as NomPE, GPE.Nom as NomGPE,');
      SQL.Add('TC.Nom as TypeC, TE.Categorie as CatE, TPE.Categorie as CatPE,');
      SQL.Add('S.Nom as NomS,S.Id_lieu as NumS');
      SQL.Add(format('from %s C join TypeCnc TC on TC.Id_typC = C.Id_typC',[NomTableCompo[eConnecteur]]));
      SQL.Add(format('join %s E on E.Id_eqt = C.Id_eqt',[NomTableCompo[eEquipement]]));
      SQL.Add('join Type TE on TE.Id_type = E.Id_type');
      SQL.Add(format('join %s PE on PE.Id_lieu = E.Id_lieu',[NomTableCompo[eGroupe]]));
      // S'il n'y a pas de GE, c'est un LT ou m�me un site dans le cas d'une terminaison
      SQL.Add('join type TPE on TPE.Id_type = PE.Id_type');
      SQL.Add(format('left join %s GPE on GPE.Id_lieu = PE.IdPere',[NomTableCompo[eLocalTechnique]]));
      // Sera inutilis� si terminaison ou pas de GE
      SQL.Add(format('join %s S on S.Id_lieu = PE.IdSite',[NomTableCompo[eLieuSimple]]));
      // Nom du site v3.5.0 (854)
      SQL.Add(format('where C.Id_cnc = %d',[peId_cnc]));
      Open;

      psTypeCnc:= FieldByName('TypeC').AsString;
      if FieldByName('CatE').AsString = InfoCateg[eEquipement].Code then  // ce n'est pas une terminaison
      begin
        psLigneeAffichee:= FlbMajusc1(InfoCateg[eEquipement].Libelle) +' : '
         + RendNonXml(FieldByName('NomE').AsString);
        if FieldByName('CatPE').AsString = InfoCateg[eGroupe].Code then
        begin
          psLigneeAffichee:= FlbMajusc1(InfoCateg[eGroupe].Libelle) +' : '
           +RendNonXml(FieldByName('NomPE').AsString)
           +sautDeLigne + psLigneeAffichee;
          nomLt:= FieldByName('NomGPE').AsString
        end
        else
          nomLt:= FieldByName('NomPE').AsString;
        psLigneeAffichee:= FlbMajusc1(InfoCateg[eLocalTechnique].Libelle) +' : '
         +RendNonXml(nomLt)    // 3.5.0 (n�???) RendNonXml oubli�
         +sautDeLigne + psLigneeAffichee;
      end
      else
        psLigneeAffichee:= FlbMajusc1(InfoCateg[eTerminaison].Libelle) +' : '
         +RendNonXml(FieldByName('NomE').AsString);

      if FieldByName('NumS').AsString <> IdSite then
        psLigneeAffichee:= DonneNomNiveauSite +' : ' +RendNonXml(FieldByName('NomS' ).AsString)
         +sautDeLigne +psLigneeAffichee;

      result:= RendNonXml(FieldByName('NomC').AsString);

    finally
      Free;
    end;
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
{ Recherche d'informations textuelles sur un connecteur � partir de son identifiant }
function clsReqGenerique.LigneeCncXml(
 peId_cnc: integer;     // Identifiant du connecteur � rechercher
 var psTailleGroupe: integer;  // structure du connecteur ou c�ble: 1=fils 2=paires 4=quartes
 var psNbGroupes: integer)   // nombre de fils,paires ou quartes du connecteur
 : string;  // renvoie une cha�ne XML contenant le LT, GE,�quipement,connecteur et connexion
{  <lt>nom_du_LT</lt>   -- absent si connecteur de terminaison
   <ge>nom_de_GE</ge>     -- absent si pas de groupe d'�quipement ou si terminaison
   <equipement>nom_de_l_equipement</equipement>
   <connecteur>nom_du_connecteur</connecteur>
   <connexion>nom_standard_de_connexion</connexion>                            }
{ ----------------------------------------------------------------------------- }

var
  CursAncetres: TADOQuery;

begin
  EntreeProc(ClassName+'.LigneeCncXml');
  CursAncetres:= CreeADOQuery();
  with CursAncetres do
    try
//      Connection := ConnexionBD;
      SQL.Add('select C.Nom as NomC, E.Nom as NomE, PE.Nom as NomPE, GPE.Nom as NomGPE,');
      SQL.Add('TC.TailleGroupe, TC.NbGroupes, TE.Categorie as CatE, TPE.Categorie as CatPE');
      SQL.Add(format('from %s C join TypeCnc TC on TC.Id_typC = C.Id_typC',[NomTableCompo[eConnecteur]]));
      SQL.Add(format('join %s E on E.Id_eqt = C.Id_eqt',[NomTableCompo[eEquipement]]));
      SQL.Add('join Type TE on TE.Id_type = E.Id_type');
      SQL.Add(format('join %s PE on PE.Id_lieu = E.Id_lieu',[NomTableCompo[eGroupe]]));
      SQL.Add('join type TPE on TPE.Id_type = PE.Id_type');
      // S'il n'y a pas de GE, c'est un LT ou m�me un site dans le cas d'une terminaison
      SQL.Add(format('join %s GPE on GPE.Id_lieu = PE.IdPere',[NomTableCompo[eLocalTechnique]]));
      // Sera inutilis� si terminaison on pas de GE
      SQL.Add(format('where C.Id_cnc =  %d',[peId_cnc]));
      Open;

      psTailleGroupe:= fieldByName('TailleGroupe').AsInteger;
      psNbGroupes:= fieldByName('NbGroupes').AsInteger;
      result:= RemplitBalise(defBalEquipement,FieldByName('NomE').AsString,true)
       +RemplitBalise(defBalConnecteur,FieldByName('NomC').AsString,true);
      if FieldByName('CatE').AsString = InfoCateg[eEquipement].Code then   // on n'est pas dans une terminaison
      begin
        if FieldByName('CatPE').AsString = InfoCateg[eGroupe].Code then
        begin
          result:= RemplitBalise(defBalLT,FieldByName('NomGPE').AsString,true)
           +RemplitBalise(defBalGE,FieldByName('NomPE').AsString,true) +result;
        end
        else
          result:= RemplitBalise(defBalLT,FieldByName('NomPE').AsString,true) + result;
      end;

    finally
      Free;
    end;
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
{ Recherche du nom d'un LT ou d'un GE
  S'il s'agit d'un GE, la proc�dure trouve aussi le nom du LT d'appartenance    }
procedure clsReqGenerique.TrouveNomLtGe(
 peId_lieu: integer;     // Identifiant de GE ou de LT
 var psNomLt: string;   // nom du LT
 var psNomGe: string);  // nom du GE (si peId_lieu correspond � un GE)
{ ----------------------------------------------------------------------------- }

var CursAncetres: TADOQuery;

begin
  EntreeProc(ClassName+'.TrouveNomLtGe');
  CursAncetres:= CreeADOQuery();
  with CursAncetres do
    try
//      Connection := ConnexionBD;
      SQL.Add('select L.Nom as NomL, LP.Nom as NomLP,T.Categorie');
      SQL.Add(format('from %s L join Type T on T.Id_type = L.Id_type',[NomTableCompo[eGroupe]]));
      // S'il n'y a pas de GE, c'est un LT ou m�me un site dans le cas d'une terminaison
      SQL.Add(format('join %s LP on LP.Id_lieu = L.IdPere',[NomTableCompo[eLocalTechnique]]));
      // Sera inutilis� si terminaison on pas de GE
      SQL.Add(format('where L.Id_lieu =  %d',[peId_lieu]));
      Open;

      if FieldByName('Categorie').AsString = InfoCateg[eGroupe].Code then  // le lieu (peId_lieu) est un GE
      begin
        psNomLt:= FieldByName('NomLP').AsString;
        psNomGe:= FieldByName('NomL').AsString;
      end
      else
      begin
        psNomLt:= FieldByName('NomL').AsString;
        psNomGe:= '';     // pas de groupe d'�quipement
      end;

    finally
      Free;
    end;
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
function clsReqGenerique.IdTypeToNom(peId_type: integer): string;
{ Donne un nom de type en fonction de son identifiant }
{ ----------------------------------------------------------------------------- }
var ADOQuery1 : TADOQuery;

begin
  EntreeProc(ClassName+'.IdTypeToNom');
  ADOQuery1 :=  CreeADOQuery();
  with ADOQuery1 do
    try
//      Connection := ConnexionBD;
      SQL.Text:= format('select Nom from Type where Id_type = %d',[peId_type]);
      Open;
      if Eof then
        result := ''
      else
        result := Fields[0].asString;
    finally
      Free;
    end;
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
{ Recherche des droits du lieu simple de n� peNumLieu et de niveau peNiveauLieu }
function clsReqGenerique.DroitAccesLieuSimple(
 peNumLieu: integer;  // num�ro du lieu � rechercher
 peNiveauLieu: integer)  // niveau du lieu
 : TDroit;
{ ----------------------------------------------------------------------------- }

var
  ind: integer;

begin
  EntreeProc(ClassName+'.DroitAccesLieuSimple');
  if ModifInhibee or ConsultationSeule then
    result:= cstDroitLectureSeule   // session d�grad�e en lecture seule
  else
    if OdtActif = 0 then // 3.4.4 (722)
    begin
      result:= 0;
      if peNiveauLieu <> 0 then
        with Query do
        begin
          SQL.Clear;
          SQL.Add('select D0.DroitLieu as Droit0');
          for ind:= 1 to peNiveauLieu do
            SQL.Add(format(',D%0:d.DroitLieu as Droit%0:d',[ind]));
          SQL.Add(format('from %s L%d',[NomTableCompo[eLieuSimple],peNiveauLieu]));
          SQL.Add(format('left join Droit_lieu D%0:d on D%0:d.Id_lieu=L%0:d.Id_lieu and D%0:d.Id_prof = %1:d',
           [peNiveauLieu,IdProfil]));
          for ind:= pred(peNiveauLieu) downto 1 do
          begin
            SQL.Add(format('join Lieu L%0:d on L%0:d.Id_lieu = L%1:d.IdPere',[ind,succ(ind)]));
            SQL.Add(format('left join Droit_lieu D%0:d on D%0:d.Id_lieu=L%0:d.Id_lieu and D%0:d.Id_prof = %1:d',
             [ind,IdProfil]));
          end;
          if peNiveauLieu > 0 then
            { Pour optimiser, on cherche le droit de la vue g�n�rale sans faire le dernier join Lieu }
            SQL.Add(format('left join Droit_lieu D0 on D0.Id_lieu=L1.IdPere and D0.Id_prof = %d',
             [IdProfil]));
          SQL.Add(format('where L%d.Id_lieu = %d',[peNiveauLieu,peNumLieu]));
          Open;
          { on remonte la hi�rarchie jusqu'� trouver une marque de droit ou jusqu'au site }
          for ind:= peNiveauLieu downto 0 do
          begin
            result:= FieldByName('Droit'+IntToStr(ind)).AsInteger;
            if result > 0 then
              BREAK;
          end;
        end;
        if result = 0 then
          result:= cstDroitModif;
    end
    else
      result:= cstDroitModif;  // 3.4.4 (722)
  SortieProc;
end;


{ ----------------------------------------------------------------------------- }
{ Recherche des droits d'un local technique ou d'un groupe d'�quipements        }
function clsReqGenerique.DroitAccesLieuTechnique(
 peNumLieu: integer;  // num�ro de LT ou de GE
 peCateg: TCategorie)   // cat�gorie de ce lieu : eLocalTechnique ou eGroupe
 : TDroit;
{ ----------------------------------------------------------------------------- }

var
  indGene,nbGenerations: integer;

begin
  EntreeProc(ClassName+'.DroitAccesLieuTechnique');
  if ModifInhibee or ConsultationSeule then
    result:= cstDroitLectureSeule   // session d�grad�e en lecture seule
  else
    if OdtActif = 0 then // 3.4.4 (722)
      with Query do
      begin
      { Recherche du lieu de n� peNumLieu, de son p�re et de son grand-p�re qui est forc�ment un LS
        pour ainsi remonter au premier lieu simple et avoir son niveau }
        if peCateg = eGroupe then
          nbGenerations:= 3
        else
          nbGenerations:= 2;
        SQL.Clear;
        SQL.Add('select T.NumNiveau,L1.IdPere');
        for indGene:= 1 to nbGenerations do
          SQL.Add(format(',D%0:d.DroitLieu as Droit%0:d',[indGene]));
        SQL.Add(format('from %s L%d',[NomTableCompo[eLieuSimple],nbGenerations]));
        if nbGenerations = 3 then
          // le p�re du lieu (numLieuGP) est donc un LT
          SQL.Add(format('join %s L2 on L2.Id_lieu = L3.IdPere',[NomTableCompo[eLocalTechnique]]));
        SQL.Add(format('join %s L1 on L1.Id_lieu = L2.IdPere',[NomTableCompo[eLieuSimple]]));

        for indGene:= 1 to nbGenerations do
          SQL.Add(format(
           'left join Droit_lieu D%0:d on D%0:d.Id_lieu = L%0:d.Id_lieu and D%0:d.Id_prof = %1:d',
           [indGene,IdProfil]));
          SQL.Add('join type T on T.Id_type = L1.Id_type');
        SQL.Add(format('where L%d.Id_lieu = %d',[nbGenerations,peNumLieu]));
        Open;

        { On essaie de trouver un droit au niveau du LT [et du GE] et du LS p�re }
        result:= 0;
        for indGene:= nbGenerations downto 1 do
        begin
          result:= FieldByName('Droit'+IntToStr(indGene)).AsInteger;
          if result > 0 then
            BREAK;
        end;
        if result = 0 then
          // pas de marque de droit jusqu'au premier LS : on va chercher les droits des LS anc�tres
          result:= DroitAccesLieuSimple(FieldByName('IdPere').AsInteger,pred(FieldByName('NumNiveau').AsInteger));
      end
    else
      result:= cstDroitModif;  // 3.4.4 (722)

  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
{ Recherche des droits et du num�ro d'un local technique ou d'un GE             }
{ La proc�dure sert aussi de test d'existence : renvoie psNumero = 0 si pas trouv� }
function clsReqGenerique.NumeroLieuTechnique(
 peNumSite: integer;   // num�ro de site concern�
 peNomSite: string;   // nom de site si recherche par nom (import - peNumSite est alors � 0)
 peNomLt: string;  // nom de LT
 peNomGe: string;  // nom de GE si c'en est un
 var psDroit: TDroit)   // droits sur ce lieu
 : integer;     // renvoie le num�ro du lieu technique cherch�
{ ----------------------------------------------------------------------------- }

var
  indGene,nbGenerations: integer;

begin
  EntreeProc(ClassName+'.NumeroLieuTechnique');
(*  v3.7.0 (1406) - remplac� par des tests (voir (1406))
  if OdtActif = 0 then // 3.4.4 (722)   *)
  with Query do
  begin
  { Recherche du lieu de n� peNumLieu, de son p�re et de son grand-p�re qui est forc�ment un LS
    pour ainsi remonter au premier lieu simple et avoir son niveau }
    if peNomGe = '' then
      nbGenerations:= 2
    else
      nbGenerations:= 3;
    SQL.Clear;
    SQL.Add(format('select T.NumNiveau,L1.IdPere,L%d.Id_lieu',[nbGenerations]));
    if not Administrateur and (OdtActif = 0) then   // test OdtActif v3.6.7e (1406)
      for indGene:= 1 to nbGenerations do
        SQL.Add(format(',D%0:d.DroitLieu as Droit%0:d',[indGene]));
    SQL.Add(format('from %s L%d',[NomTableCompo[eLieuSimple],nbGenerations]));
    if nbGenerations = 3 then
      // le p�re du lieu cherch� est donc un LT
      SQL.Add(format('join %s L2 on L2.Id_lieu = L3.IdPere',[NomTableCompo[eLocalTechnique]]));
    SQL.Add(format('join %s L1 on L1.Id_lieu = L2.IdPere',[NomTableCompo[eLieuSimple]]));

    if peNumSite = 0 then    // Si recherche par nom
      SQL.Add(format('join %s S on S.Id_lieu = L1.IdSite',[NomTableCompo[eLieuSimple]]));
    if not Administrateur and (OdtActif = 0) then   // test OdtActif v3.6.7e (1406)
      for indGene:= 1 to nbGenerations do
        SQL.Add(format(
         'left join Droit_lieu D%0:d on D%0:d.Id_lieu = L%0:d.Id_lieu and D%0:d.Id_prof = %1:d',
         [indGene,IdProfil]));
    SQL.Add('join type T on T.Id_type = L1.Id_type');

    if peNumSite = 0 then    // Si recherche par nom
      SQL.Add(format('where S.Nom = %s',[quotedStr(peNomSite)]))
    else
      SQL.Add(format('where L%d.IdSite = %d',[nbGenerations,peNumSite]));
    if peNomGe = '' then  // il n'y a qu'un LT
      SQL.Add(format('and L%d.Nom = ''%s''',[nbGenerations,DoubleQuotes(peNomLt)]))
    else
      SQL.Add(format('and L%d.Nom = ''%s'' and L%d.Nom = ''%s''',
       [nbGenerations,DoubleQuotes(peNomGe),pred(nbGenerations),DoubleQuotes(peNomLt)]));
    Open;
    if Eof then
      result:= 0
    else
    begin
      result:= fieldByName('Id_lieu').asInteger;
      { On essaie de trouver un droit au niveau du LT [et du GE] et du LS p�re }
      if Administrateur or (OdtActif > 0) then  // test OdtActif v3.6.7e (1406)
        psDroit:= cstDroitModif
      else
      begin
        indGene:= nbGenerations;
        while (psDroit = 0) and (indGene > 0) do
        begin
          psDroit:= FieldByName('Droit'+IntToStr(indGene)).AsInteger;
          dec(indGene)
        end;
        if psDroit = 0 then
          // pas de marque de droit jusqu'au premier LS : on va chercher les droits des LS anc�tres
          psDroit:= DroitAccesLieuSimple(FieldByName('IdPere').AsInteger,pred(FieldByName('NumNiveau').AsInteger));
      end;
      if (ModifInhibee or ConsultationSeule) and (psDroit <= cstDroitModif) then
        psDroit:= cstDroitLectureSeule;   // si psDroit = cstDroitAucunAcces, on le laisse tel quel
    end;
  end;
(*  supprim� v3.7.0 (1406)  else  // v3.4.4
    result:= cstDroitModif;    le r�sultat ne doit pas �tre un droit mais un Id_lieu !!!
*)
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
{ Recherche des droits de l'�quipement ou de la terminaison de n� peNumEqt }
function clsReqGenerique.DroitAccesEquipement(
 peNumEqt: integer)  // num�ro de l'�quipement � rechercher
 : TDroit;
{ ----------------------------------------------------------------------------- }

var
  indGene,numLieuGP,nbGenerations: integer;
  categPere: char;

begin
  EntreeProc(ClassName+'.DroitAccesEquipement');
  if OdtActif = 0 then // 3.4.4 (722)
    with Query do
    begin
      { 1) Recherche de l'�quipement, de son p�re, de la cat�gorie et du droit de son p�re }
      SQL.Clear;
      SQL.Add(format('select IdPere,Categorie,DroitLieu,NumNiveau from %s E',
       [NomTableCompo[eEquipement]]));
      SQL.Add(format('join %s L on L.Id_lieu = E.Id_lieu',[NomTableCompo[eLieuSimple]]));
      SQL.Add(format('left join Droit_lieu D on D.Id_lieu = L.Id_lieu and D.Id_prof = %d',[IdProfil]));
      SQL.Add('join Type T on T.Id_type=L.Id_type');
      SQL.Add(format('where E.Id_eqt = %d',[peNumEqt]));
      Open;

      categPere:= FieldByName('Categorie').AsString[1];
      numLieuGP:= FieldByName('IdPere').AsInteger;
      result:= FieldByName('DroitLieu').AsInteger;

      if result = 0 then   // pas trouv� de droit
        if categPere = InfoCateg[eLieuSimple].Code then  // cas des terminaisons les plus courantes
          result:= DroitAccesLieuSimple(numLieuGP,pred(fieldByName('NumNiveau').asInteger))
        else   // categPere = LT ou GP

        { 2) Recherche du lieu de n� numLieuGP et si c'est un LT, de son p�re qui est forc�ment un LS
             pour ainsi remonter au premier lieu simple et avoir son niveau }
        begin
          if categPere = InfoCateg[eGroupe].Code then
            nbGenerations:= 2
          else
            nbGenerations:= 1;
          SQL.Clear;
          SQL.Add('select T.NumNiveau,L1.IdPere');
          for indGene:= 1 to nbGenerations do
            SQL.Add(format(',D%0:d.DroitLieu as Droit%0:d',[indGene]));
          SQL.Add(format('from %s L%d',[NomTableCompo[eLieuSimple],nbGenerations]));
          if nbGenerations = 2 then  // <=> categPere = InfoCateg[eGroupe].Code
            // le grand-p�re de l'eqt (numLieuGP) est donc un LT
            SQL.Add(format('join %s L1 on L1.Id_lieu = L2.IdPere',[NomTableCompo[eLieuSimple]]));
          for indGene:= 1 to nbGenerations do
            SQL.Add(format(
             'left join Droit_lieu D%0:d on D%0:d.Id_lieu = L%0:d.Id_lieu and D%0:d.Id_prof = %1:d',
             [indGene,IdProfil]));
          SQL.Add('join type T on T.Id_type = L1.Id_type');
          SQL.Add(format('where L%d.Id_lieu = %d',[nbGenerations,numLieuGP]));
          Open;

          { On essaie de trouver un droit au niveau du LT [et du GE] et du LS p�re }
          indGene:= nbGenerations;
          while (result = 0) and (indGene > 0) do
          begin
            result:= FieldByName('Droit'+IntToStr(indGene)).AsInteger;
            dec(indGene)
          end;
          if result = 0 then
            // pas de marque de droit jusqu'au premier LS : on va chercher les droits des LS anc�tres
            result:= DroitAccesLieuSimple(FieldByName('IdPere').AsInteger,pred(FieldByName('NumNiveau').AsInteger));
        end;

      if (ModifInhibee or ConsultationSeule) and (result <= cstDroitModif) then
        result:= cstDroitLectureSeule;   // si psDroit = cstDroitAucunAcces, on le laisse tel quel
    end
  else  // v3.4.4
    result:= cstDroitModif;

  SortieProc;
end;

{ --------------------------------------------------------------------------------------------- }
procedure clsreqGenerique.CreerTablesTempo;
{ Cr�ation de tables temporaires en copiant les tables originales - la requ�te appelante va
  utiliser ces tables (pour un filtre ou une simulation)
  Contexte : initialement, les noms de tables ont la valeur standard (pas de filtre en place)   }
{ --------------------------------------------------------------------------------------------- }

var
  debutNom: string;
  indCat: TCategorie;
  indAutre: TAutreTable;

begin
  EntreeProc(ClassName+'.CreerTablesTempo');
  debutNom:= 'TMP'+CompleteAZero(NumSession,10);
  { 1) Changement de nom dans les 2 tableaux de noms de tables }
  for indCat:= low(TCategorie) to high(TCategorie) do
    NomTableCompo[indCat]:= debutNom+NomTableCompo[indCat];
  for indAutre:= low(TAutreTable) to high(TAutreTable) do
    NomAutreTable[indAutre]:= debutNom+NomAutreTable[indAutre];

  { 2) Cr�ation par select * de toutes les copies de tables en utilisant les nouveaux noms }
  with Query do
  begin
    SQL.Clear;
    if Idsite = '' then
    begin
      SQL.Add(format('select * into %s from %s',
       [NomTableCompo[eLieuSimple],NomTableReelle(eLieuSimple)]));
      SQL.Add(format('select * into %s from %s',
       [NomTableCompo[eEquipement],NomTableReelle(eEquipement)]));
      SQL.Add(format('select * into %s from %s',
       [NomTableCompo[eConnecteur],NomTableReelle(eConnecteur)]));
      SQL.Add(format('select * into %s from %s',
       [NomTableCompo[eFonction],NomTableReelle(eFonction)]));
      SQL.Add(format('select * into %s from %s',
       [NomTableCompo[eCable],NomTableReelle(eCable)]));
      SQL.Add(format('select * into %s from Lien',[NomAutreTable[eLien]]));
      SQL.Add(format('select * into %s from Parcours',[NomAutreTable[eParcours]]));
      SQL.Add(format('select * into %s from Depart',[NomAutreTable[eDepart]]));
      SQL.Add(format('select * into %s from Cablage',[NomAutreTable[eCablage]]));
      SQL.Add(format('select * into %s from Symbole',[NomAutreTable[eSymbole]]));
      SQL.Add(format('select * into %s from ValeurProp',[NomAutreTable[eValeurProp]]));
      { Nouvelles tables 3.6.0 }
      SQL.Add(format('select * into %s from PlanLieu',[NomTableCompo[ePlanLieu]]));
      SQL.Add(format('select * into %s from Trajet',[NomAutreTable[eTrajet]]));
      SQL.Add(format('select * into %s from ExtTrajet',[NomAutreTable[eExtTrajet]]));
      SQL.Add(format('select * into %s from Trajet_famille',[NomAutreTable[eTrajet_famille]]));
    end
    else
    begin
      { Lieux appartenant au site ou avec IdSite � NULL (mod�les, �l�ments pr�cr��s)}
      SQL.Add(format('select * into %s from %s',
       [NomTableCompo[eLieuSimple],NomTableReelle(eLieuSimple)]));
      SQL.Add(format('where IdSite = %s or IdSite is null or IdPere is null',[IdSite]));
      // On inclut la vue g�n�rale (IdPere is null)

      { Equipements et terminaisons appartenant � ces lieux ou eqts mod�les ou pr�cr��s }
      SQL.Add(format('select * into %s from %s',
       [NomTableCompo[eEquipement],NomTableReelle(eEquipement)]));
      { Attention: pas de jointure sinon on perd la propri�t� d'identit� de Id_eqt }
      SQL.Add(format('where Id_lieu is null or Id_lieu in (select Id_lieu from %s)',
       [NomTableCompo[eLocalTechnique]]));

      { Connecteurs appartenant � ces �quipements ou pr�cr��s par ODT }
      SQL.Add(format('select * into %s from %s',
       [NomTableCompo[eConnecteur],NomTableReelle(eConnecteur)]));
      { Attention: pas de jointure sinon on perd la propri�t� d'identit� de Id_cnc }
      SQL.Add(format('where Id_eqt is null or Id_eqt in (select Id_eqt from %s)',
       [NomTableCompo[eEquipement]]));

      { Fonctions appartenant au site ou mod�les ou pr�cr��s }
      SQL.Add(format('select * into %s from %s',
       [NomTableCompo[eFonction],NomTableReelle(eFonction)]));
      SQL.Add(format('where IdSite = %s or IdSite is null',[IdSite]));

      { C�bles et bo�tes noires appartenant au site ou mod�les ou pr�cr��s }
      SQL.Add(format('select * into %s from %s',
       [NomTableCompo[eCable],NomTableReelle(eCable)]));
      SQL.Add(format('where IdSite = %s or IdSite is null',[IdSite]));

      { D�parts sur des connecteurs de la table temporaire }
      SQL.Add(format('select D.* into %s from Depart D',[NomAutreTable[eDepart]]));
      SQL.Add(format('join %s C on C.Id_cnc=D.Id_cnc',[NomTableCompo[eConnecteur]]));

      { C�blages appartenant aux connecteurs de la table temporaire }
      SQL.Add(format('select Cb.* into %s from Cablage Cb',[NomAutreTable[eCablage]]));
      SQL.Add(format('join %s C on C.Id_cnc=Cb.Id_cnc',[NomTableCompo[eConnecteur]]));

      { Liens d'appartenance de ces c�blages }
      SQL.Add(format('select * into %s from Lien L',[NomAutreTable[eLien]]));
      { Attention: pas de jointure sinon on perd la propri�t� d'identit� de Id_lien }
      SQL.Add(format('where Id_lien in (select Id_lien from %s)',[NomAutreTable[eCablage]]));

      { Parcours d'appartenance de ces liens }
      SQL.Add(format('select * into %s from Parcours',[NomAutreTable[eParcours]]));
      { Attention: pas de jointure sinon on perd la propri�t� d'identit� de Id_par }
      SQL.Add(format('where Id_par in (select Id_par from %s)',[NomAutreTable[eLien]]));

      { Nouvelles tables 3.6.0 }
      { Plans de lieux (la table doit �tre cr��e avant la requ�te des symboles) }
      SQL.Add(format('select * into %s from %s',[NomTableCompo[ePlanLieu],InfoTable[eTablePlanLieu].NomTableReelle]));
      { Attention: pas de jointure sinon on perd la propri�t� d'identit� de Id_plan }
      SQL.Add(format('where Id_lieu is null or Id_lieu in (select Id_lieu from %s)',[NomTableCompo[eLieuSimple]]));

      { Trajets sur ces plans }
      SQL.Add(format('select * into %s from Trajet',[NomAutreTable[eTrajet],InfoTable[eTableTrajet].NomTableReelle]));
      { Attention: pas de jointure sinon on perd la propri�t� d'identit� de Id_plan }
      SQL.Add(format('where Id_plan is null or Id_plan in (select Id_plan from %s)',[NomTableCompo[ePlanLieu]]));

      { Extr�mit�s de ces trajets }
      SQL.Add(format('select ET.* into %s from ExtTrajet ET',[NomAutreTable[eExtTrajet]]));
      SQL.Add(format('join %s T on T.Id_traj=ET.Id_traj',[NomAutreTable[eTrajet]]));

      { Famille d'appartenance de ces trajets }
      SQL.Add(format('select TF.* into %s from Trajet_Famille TF',[NomAutreTable[eTrajet_Famille]]));
      SQL.Add(format('join %s T on T.Id_traj=TF.Id_traj',[NomAutreTable[eTrajet]]));

      { Symboles appartenant � des plans de de lieux ou d'�quipements des tables temporaires }
      SQL.Add(format('select S.* into %s from Symbole S',[NomAutreTable[eSymbole]]));
      SQL.Add(format('left join %s P on P.Id_plan=S.Id_plan',[NomTableCompo[ePlanLieu]]));    // v3.6.0
      SQL.Add(format('left join %s E on E.Id_eqt=S.Id_eqt',[NomTableCompo[eEquipement]]));
      SQL.Add('where P.Id_plan is not null or E.Id_eqt is not null');  // une des deux jointures doit avoir "march�"

      { Valeurs de propri�t�s }
      SQL.Add(format('select VP.* into %s from ValeurProp VP',[NomAutreTable[eValeurProp]]));
      SQL.Add(format('left join %s L on VP.Categorie in (''%s'',''%s'',''%s'') and L.Id_lieu=VP.IdObjet',
       [NomTableCompo[eLieuSimple],InfoCateg[eLieuSimple].Code,InfoCateg[eLocalTechnique].Code,InfoCateg[eGroupe].Code]));
      SQL.Add(format('left join %s E on VP.Categorie in (''%s'',''%s'') and E.Id_eqt=VP.IdObjet',
       [NomTableCompo[eEquipement],InfoCateg[eEquipement].Code,InfoCateg[eTerminaison].Code]));
      SQL.Add(format('left join %s C on VP.Categorie in (''%s'',''%s'') and C.Id_cab=VP.IdObjet',
       [NomTableCompo[eCable],InfoCateg[eBoiteNoire].Code,InfoCateg[eCable].Code]));
      SQL.Add(format('left join %s F on VP.Categorie = ''%s'' and F.Id_fon=VP.IdObjet',
       [NomTableCompo[eFonction],InfoCateg[eFonction].Code]));
      { v3.5.4: propri�t�s de brassage (liens) }
      SQL.Add(format('left join %s Z on VP.Categorie = ''%s'' and Z.Id_lien=VP.IdObjet',
       [NomTableCompo[eBrassage],infoCateg[eBrassage].Code]));

      SQL.Add(format('where Categorie = ''%s'' or L.Id_lieu is not null or E.Id_eqt is not null',
       [InfoCateg[eOrdreTravaux].Code]));   // inclut les propri�t�s des ODT (tous)
      SQL.Add('or C.Id_cab is not null or F.Id_fon is not null or Z.Id_lien is not null');  // v3.6.6 (1315)
    end;
    SQL.Add('select 0 [Resultat]');    // pour que la requ�te renvoie un r�sultat

    Open;

    { Cr�ation des m�mes index que les tables r�elles (essai r�solution bug 1200) }
    SQL.Clear;
    SQL.Add(format('create index %sLieuIIdPere on %s (IdPere)',
     [debutNom,NomTableCompo[eLieuSimple]]));
    SQL.Add(format('create index %sLieuIIdSite on %s (IdSite)',
     [debutNom,NomTableCompo[eLieuSimple]]));
    SQL.Add(format('create index %sEquipementIId_Lieu on %s (Id_lieu)',
     [debutNom,NomTableCompo[eEquipement]]));
    SQL.Add(format('create index %sConnecteurIId_eqt on %s (Id_eqt)',
     [debutNom,NomTableCompo[eConnecteur]]));
    SQL.Add(format('create index %sCablageIId_cnc on %s (Id_cnc)',
     [debutNom,NomAutreTable[eCablage]]));
    SQL.Add(format('create index %sCablageIId_lien on %s (Id_lien)',
     [debutNom,NomAutreTable[eCablage]]));
    SQL.Add(format('create index %sLienIId_cab on %s (Id_cab)',
     [debutNom,NomAutreTable[eLien]]));
    SQL.Add(format('create index %sLienIId_par on %s (Id_par)',
     [debutNom,NomAutreTable[eLien]]));
    SQL.Add(format('create index %sParcoursIId_fon on %s (Id_fon)',
     [debutNom,NomAutreTable[eParcours]]));
    SQL.Add(format('create index %sDepartIId_cnc on %s (Id_cnc)',
     [debutNom,NomAutreTable[eDepart]]));
    SQL.Add(format('create index %sDepartIId_fon on %s (Id_fon)',
     [debutNom,NomAutreTable[eDepart]]));
    SQL.Add(format('create index %sFonctionIIdSite on %s (IdSite)',
     [debutNom,NomTableCompo[eFonction]]));
    SQL.Add(format('create index %sCableIIdSite on %s (IdSite)',
     [debutNom,NomTableCompo[eCable]]));
    SQL.Add(format('create index %sSymboleIId_plan on %s (Id_plan)',
     [debutNom,NomAutreTable[eSymbole]]));
    SQL.Add(format('create index %sSymboleIId_eqt on %s (Id_eqt)',
     [debutNom,NomAutreTable[eSymbole]]));
    SQL.Add(format('create index %sValeurPropIIdProp on %s (Id_prop)',
     [debutNom,NomAutreTable[eValeurProp]]));
    SQL.Add(format('create index %sValeurPropIIdObjet on %s (IdObjet)',
     [debutNom,NomAutreTable[eValeurProp]]));
    { index sur les nouvelles tables 3.6.0 li�s aux FK }
    SQL.Add(format('create index %sPlanLieuIId_lieu on %s(Id_lieu)',
     [debutNom,NomTableCompo[ePlanLieu]]));
    SQL.Add(format('create index %sTrajetIId_plan on %s(Id_plan)',
     [debutNom,NomAutreTable[eTrajet]]));
    SQL.Add(format('create index %sExtTrajetIId_traj on %s(Id_traj)',
     [debutNom,NomAutreTable[eExtTrajet]]));
    SQL.Add(format('create index %sTrajet_familleIId_traj on %s(Id_traj)',
     [debutNom,NomAutreTable[eTrajet_famille]]));
    SQL.Add(format('create index %sTrajet_familleIId_fam on %s(Id_fam)',
     [debutNom,NomAutreTable[eTrajet_famille]]));

    { v3.5.5a: ajout d'index sur tous les champs qui sont des PK dans les tables d'origine
      (un test a montr� une acc�l�ration de la simulation des actions d'un ODT) }
    SQL.Add(format('create index %sLieuIId_lieu on %s (Id_lieu)',
     [debutNom,NomTableCompo[eLieuSimple]]));
    SQL.Add(format('create index %sEquipementIId_par on %s (Id_eqt)',
     [debutNom,NomTableCompo[eEquipement]]));
    SQL.Add(format('create index %sConnecteurIId_cnc on %s (Id_cnc)',
     [debutNom,NomTableCompo[eConnecteur]]));
    SQL.Add(format('create index %sFonctionIId_fon on %s (Id_fon)',
     [debutNom,NomTableCompo[eFonction]]));
    SQL.Add(format('create index %sCableIId_cab on %s (Id_cab)',
     [debutNom,NomTableCompo[eCable]]));
    SQL.Add(format('create index %sLienIId_lien on %s (Id_lien)',
     [debutNom,NomAutreTable[eLien]]));
    SQL.Add(format('create index %sParcoursIId_par on %s (Id_par)',
     [debutNom,NomAutreTable[eParcours]]));
    SQL.Add(format('create index %sPlanLieuIId_plan on %s (Id_plan)',
     [debutNom,NomTableCompo[ePlanLieu]]));
    SQL.Add(format('create index %sTrajetIId_traj on %s (Id_traj)',
     [debutNom,NomAutreTable[eTrajet]]));
    ExecSQL;    // et pas Open car la commande ne renvoie pas de r�sultat

    { v3.5.0 (854) on inclut les c�blages qui m�nent dans un autre site (initialement non pr�sents dans la table temporaire) }
    SQL.Clear;
    SQL.Add(format('insert into %s (Id_cnc,Id_lien,Cnx,Origine)',[NomAutreTable[eCablage]]));
    SQL.Add('select CbR.Id_cnc,CbR.Id_lien,CbR.Cnx,CbR.Origine from Cablage CbR');
    { Prend seulement les c�blages de la table r�elle dont la r�f�rence de lien est dans la table temporaire des liens}
    SQL.Add(format('join %s Ln on Ln.Id_lien = CbR.Id_lien',[NomAutreTable[eLien]]));
    { Mais qui r�f�rencent un connecteur inexistant dans la table temporaire }
    SQL.Add(format('where not exists (select CT.Id_cnc from %s CT where CT.Id_cnc = CbR.Id_cnc)',[NomTableCompo[eConnecteur]]));
    ExecSQL;

    if RowsAffected > 0 then   // Sinon il n'y a aucun lien dans le site de l'ODT qui part vers un autre site
    begin
      SQL.Clear;
      { v3.5.0 (854) on inclut les c�bles et bo�tes noires (intersites) de ces liens qui ne sont pas encore dans la table temporaire }
      SQL.Add(format('set identity_insert %s on',[NomTableCompo[eCable]]));
      SQL.Add(format('insert into %s (Id_cab,IdSite,Id_type,Nom,TailleGroupe,NbGroupes,Longueur,Comm,Id_act)',
       [NomTableCompo[eCable]]));
      SQL.Add(
       'select distinct CR.Id_cab,CR.IdSite,CR.Id_type,CR.Nom,CR.TailleGroupe,CR.NbGroupes,CR.Longueur,CR.Comm,CR.Id_act');
      SQL.Add(format('from %s CR',[NomTableReelle(eCable)]));
      { Prend seulement les c�blages de la table r�elle dont la r�f�rence de lien est dans la table temporaire des liens}
      SQL.Add(format('join %s Ln on Ln.Id_cab = CR.Id_cab',[NomAutreTable[eLien]]));
      { Mais qui r�f�rencent un c�ble inexistant dans la table temporaire }
      SQL.Add(format('where not exists (select CT.Id_cab from %s CT where CT.Id_cab = CR.Id_cab)',
      [NomTableCompo[eCable]]));
      SQL.Add(format('set identity_insert %s off',[NomTableCompo[eCable]]));

      { v3.5.0 (854) on inclut les connecteurs de ces c�blages qui ne sont pas encore dans la table temporaire }
      SQL.Add(format('set identity_insert %s on',[NomTableCompo[eConnecteur]]));
      SQL.Add(format('insert into %s (Id_cnc,Id_typC,Id_eqt,Nom,Classement,Finligne,Id_act)',[NomTableCompo[eConnecteur]]));
      SQL.Add(format('select distinct CR.Id_cnc,CR.Id_typC,CR.Id_eqt,CR.Nom,CR.Classement,CR.Finligne,CR.Id_act from %s CR',
       [NomTableReelle(eConnecteur)]));
      { Prend seulement les connecteurs de la table r�elle qui ont un c�blage dans la table temporaire }
      SQL.Add(format('join %s Cb on Cb.Id_cnc = CR.Id_cnc',[NomAutreTable[eCablage]]));
      { Mais qui n'existent pas dans la table temporaire }
      SQL.Add(format('where not exists (select CT.Id_cnc from %s CT where CT.Id_cnc = CR.Id_cnc)',
       [NomTableCompo[eConnecteur]]));
      SQL.Add(format('set identity_insert %s off',[NomTableCompo[eConnecteur]]));

      { v3.5.0 (854) on inclut les �quipements de ces connecteurs }
      SQL.Add(format('set identity_insert %s on',[NomTableCompo[eEquipement]]));
      SQL.Add(format(
       'insert into %s (Id_eqt,Id_lieu,Id_type,Nom,Role,Comm,Id_act,NomFichier,Extension,LargeurPlan)',
       [NomTableCompo[eEquipement]]));
      SQL.Add(format(
       'select distinct ER.Id_eqt,ER.Id_lieu,ER.Id_type,ER.Nom,ER.Role,ER.Comm,ER.Id_act,ER.NomFichier,ER.Extension,ER.LargeurPlan from %s ER',
       [NomTableReelle(eEquipement)]));
      { Prend seulement les �quipements de la table r�elle qui ont un connecteur dans la table temporaire }
      SQL.Add(format('join %s Cn on Cn.Id_eqt = ER.Id_eqt',[NomTableCompo[eConnecteur]]));
      { Mais qui n'existent pas dans la table temporaire }
      SQL.Add(format('where not exists (select ET.Id_eqt from %s ET where ET.Id_eqt = ER.Id_eqt)',
       [NomTableCompo[eEquipement]]));;
      SQL.Add(format('set identity_insert %s off',[NomTableCompo[eEquipement]]));

      { v3.5.0 (854) on inclut les lieux d'appartenance de ces �quipements (GE, LT ou lieu simple pour les terminaisons) }
      SQL.Add(format('set identity_insert %s on',[NomTableCompo[eGroupe]]));
      SQL.Add(format('insert into %s (Id_lieu,IdPere,Id_type,IdSite,Nom,Classement,Comm,Id_act,Ordre)',
       [NomTableCompo[eGroupe]]));    // eGroupe pour se souvenir que �a peut commencer au niveau GE mais on peut prendre eLocalTechnique ou eLieuSimple
      SQL.Add(format('select distinct LR.Id_lieu,LR.IdPere,LR.Id_type,LR.IdSite,LR.Nom,LR.Classement,LR.Comm,LR.Id_act,LR.Ordre from %s LR',
       [NomTableReelle(eGroupe)]));
      { Prend seulement les lieux de la table r�elle qui ont un �quipement dans la table temporaire }
      SQL.Add(format('join %s E on E.Id_Lieu = LR.Id_lieu',[NomTableCompo[eEquipement]]));
      { Mais qui n'existent pas dans la table temporaire }
      SQL.Add(format('where not exists (select LT.Id_lieu from %s LT where LT.Id_Lieu = LR.Id_lieu)',
       [NomTableCompo[eGroupe]]));;
      SQL.Add(format('set identity_insert %s off',[NomTableCompo[eGroupe]]));

      ExecSQL;

      { Remonter jusqu'au niveau site en prenant les lieux p�res (la vue g�n�rale est d�j� en table temporaire }
      { v3.5.0 (854) on inclut les lieux d'appartenance de ces �quipements (GE, LT ou lieu simple pour les terminaisons) }
      SQL.Clear;
      SQL.Add(format('set identity_insert %s on',[NomTableCompo[eLieuSimple]]));
      ExecSQL;

      SQL.Clear;
      SQL.Add(format(
       'insert into %s (Id_lieu,IdPere,Id_type,IdSite,Nom,Classement,Comm,Id_act,Ordre)',
       [NomTableCompo[eLieuSimple]]));
      SQL.Add(format(
       'select distinct LR.Id_lieu,LR.IdPere,LR.Id_type,LR.IdSite,LR.Nom,LR.Classement,LR.Comm,LR.Id_act,LR.Ordre from %s LR',
       [NomTableReelle(eLieuSimple)]));
      { Prend seulement les lieux de la table r�elle qui ont un lieu fils dans la table temporaire }
      SQL.Add(format('join %s LF on LF.IdPere = LR.Id_lieu',[NomTableCompo[eLieuSimple]]));
      { Mais qui n'existent pas dans la table temporaire }
      SQL.Add(format('where not exists (select LT.Id_lieu from %s LT where LT.Id_Lieu = LR.Id_lieu)',
       [NomTableCompo[eLieuSimple]]));
  //    SQL.Add(format('set identity_insert %s off',[NomTableCompo[eGroupe]]));

      repeat
        ExecSQL;
      until RowsAffected = 0;   // Attention : RowsAffected concerne la derni�re requ�te: il faut que ce soit le insert et pas set identity_insert off

      { Remettre l'identity insert de la table temporaire des Lieux � off }
      SQL.Clear;
      SQL.Add(format('set identity_insert %s off',[NomTableCompo[eLieuSimple]]));
      ExecSQL;
    end;   // if RowsAffected > 0

  end;
  SortieProc;
end;

{ --------------------------------------------------------------------------------------------- }
procedure clsReqGenerique.SupprimerTablesTempo;
{ Suppression des tables temporaires li�es � la session (NumSession) et � la vue (NumVue)       }
{ Contexte: toutes les tables temporaires existent                                              }
{ --------------------------------------------------------------------------------------------- }
begin
  EntreeProc(ClassName+'.SupprimerTablesTempo');
  with Query do
  begin
    SQL.Clear;
    SQL.Add(format('drop table %s',[NomTableCompo[eLieuSimple]]));
    SQL.Add(format('drop table %s',[NomTableCompo[eEquipement]]));
    SQL.Add(format('drop table %s',[NomTableCompo[eConnecteur]]));
    SQL.Add(format('drop table %s',[NomTableCompo[eFonction]]));
    SQL.Add(format('drop table %s',[NomTableCompo[eCable]]));
    SQL.Add(format('drop table %s',[NomTableCompo[ePlanLieu]]));    // v3.6.0
    SQL.Add(format('drop table %s',[NomAutreTable[eLien]]));
    SQL.Add(format('drop table %s',[NomAutreTable[eParcours]]));
    SQL.Add(format('drop table %s',[NomAutreTable[eDepart]]));
    SQL.Add(format('drop table %s',[NomAutreTable[eCablage]]));
    SQL.Add(format('drop table %s',[NomAutreTable[eSymbole]]));
    SQL.Add(format('drop table %s',[NomAutreTable[eValeurProp]]));
    SQL.Add(format('drop table %s',[NomAutreTable[eTrajet]]));        // v3.6.0
    SQL.Add(format('drop table %s',[NomAutreTable[eExtTrajet]]));        // v3.6.0
    SQL.Add(format('drop table %s',[NomAutreTable[eTrajet_famille]]));     // v3.6.0
    ExecSQL;
  end;
  SortieProc;
end;


{ ---------------------------------------------------------------------------- }
{ Formatage d'une ou plusieurs connexions ou d'un ou plusieurs fils de c�ble pour affichage }
function clsReqGenerique.NomCnx(
 peCnx: integer;  // R�f�rence de connexion (commen�ant � 1)
 peNbCnx: integer;   // Nombre de connexions � regrouper sous la m�me appellation si possible
 peTailleGroupe: integer;    // Taille du groupe dans le connecteur ou le c�ble
 peNbGroupes: integer;      // Nombre de groupes dans le connecteur ou le c�ble
 var psGroupagePossible: boolean;   // Vrai si toutes les (peNbCnx) connexions ont pu �tre group�es sous ce nom, faux sinon
 peRetourVidePossible: boolean = true)  // Vrai s'il faut renvoyer '' quand le nombre de Cnx correspond � celui du connecteur ou du c�ble
 : string;   // Renvoie la cha�ne pr�te � afficher
{ ---------------------------------------------------------------------------- }

begin
  EntreeProc(ClassName+'.NomCnx');
  if peRetourVidePossible and (peNbCnx = peTailleGroupe*peNbGroupes) and (peCnx = 1) then
  begin
    psGroupagePossible:= true;   // La s�rie de cnx couvre tout le connecteur
    result:= ''
  end
  else
    if (peNbCnx = peTailleGroupe) and (pred(peCnx) mod peTailleGroupe = 0) then
    begin
      psGroupagePossible:= true;
      // Les peNbCnx connexions correspondent � un groupe: on retourne le nom de groupe
      case peTailleGroupe of
        1: result:= intToStr(peCnx);
        2: result:= 'P'+intToStr(succ(pred(peCnx) div 2));
        4: result:= 'Q'+intToStr(succ(pred(peCnx) div 4));
      end
    end
    else
    begin
      psGroupagePossible:= false;
      if peTailleGroupe = 1 then
        result:= intToStr(peCnx)
      else
      begin
        if peNbGroupes = 1 then
          result:= ''
        else
          result:= intToStr(succ(pred(peCnx) div peTailleGroupe));
        result:=  result + chr(65 + pred(peCnx) mod peTailleGroupe)
      end;
    end;
  SortieProc;
end;


{ ---------------------------------------------------------------------------- }
{ D�signation standard d'une s�rie de connexions ou de fils de c�ble pour affichage }
function clsReqGenerique.LibelleSerieCnxOuFils(
 pePremierFil: integer;    // premier fil ou connexion
 peNbCnx: integer;         // nombre de connexions de la s�rie
 peTailleGroupe: integer;  // taille d'un groupe = 1 (fils), 2 (paires) ou 4 (quartes)
 peNbGroupes: integer) // nombre de groupes (en fait, on teste s'il est = 1 ou pas)
 : string;    // Renvoie la cha�ne au format standard
 // cette cha�ne sera libell�e sous forme d'une ou deux plages selon le cas
 // (ex: "1A", "P2", "P1 � P4", "2D � Q3", "1B � 2A", "2 � 5")

var
  dernierFil,finPlage1,debutPlage2: integer;
  aFournir: boolean;

begin
  EntreeProc(ClassName+'.LibelleSerieCnxOuFils');
  dernierFil:= pred(pePremierFil+peNbCnx);

  if (pePremierFil mod peTailleGroupe = 1)  // premierFil au d�but d'un groupe
   and (dernierFil >= pePremierFil+pred(peTailleGroupe)) then // et dernierFil � la fin du groupe ou au del�
  	finPlage1:= pePremierFil+pred(peTailleGroupe)
  else
  	finPlage1:= pePremierFil;  //  impossible de simplifier la plage

  if (dernierFil mod peTailleGroupe = 0)  // Si la plage s'arr�te � la fin d'un groupe
   and (pePremierFil <= dernierFil-pred(peTailleGroupe)) then
    debutPlage2:= dernierFil-pred(peTailleGroupe)
  else
    debutPlage2:= dernierFil;

  if (finPlage1 = dernierFil) and (debutPlage2 = pePremierFil) then
		// tout peut �tre regroup� dans une seule d�signation
	  result:= NomCnx(pePremierFil,peNbCnx,peTailleGroupe,peNbGroupes,aFournir,false)
  else
  	result:= NomCnx(pePremierFil,succ(finPlage1-pePremierFil),peTailleGroupe,peNbGroupes,aFournir,false)
	   + ' '+lbA+' '
     + NomCnx(debutPlage2,succ(dernierFil-debutPlage2),peTailleGroupe,peNbGroupes,aFournir,false);
  SortieProc;
end;

{ --------------------------------------------------------------------------------------------- }
function clsReqGenerique.LitRegistreADN(
 peNomCle: string;   // nom de la cl� � lire dans
 peNomValeur: string;   // nom de la valeur � lire dans cette cl�
 var psValeurLue: string)   // valeur lue renvoy�e � l'appelant
 : boolean;
{ --------------------------------------------------------------------------------------------- }

begin
  EntreeProc(ClassName+'.LitRegistreADN');
  result:= LitRegistres(cstRegR3Web+Environnement+'\'+peNomCle,peNomValeur,psValeurLue);
  SortieProc;
end;


{ --------------------------------------------------------------------------------------------- }
{ transform�e v3.6.5a }
function clsReqGenerique.DonneCheminAcces(
 peNomValeurRegistre: string;   // = cstRegR3Serveur ou cstRegPlans ou cstRegSymboles
 peCheminComplet: boolean = true)   // false s'il faut juste donner le nom du sous-dossier pour cstRegPlans ou cstRegSymboles
 : string;     // renvoie le chemin d'acc�s lu dans la base de registres correspondant au param�tre fourni
{ --------------------------------------------------------------------------------------------- }
var
  cheminAccesLu: string;

begin
  EntreeProc(ClassName+'.DonneCheminAcces');

  if peCheminComplet or (peNomValeurRegistre = cstRegR3Serveur) then
  //  si peNomValeurRegistre = cstRegR3Serveur, on renvoie le chemin complet car il est �crit ainsi dans la base de registres
    if LitRegistreADN(cstRegCheminsDAcces,cstRegR3Serveur,cheminAccesLu)
     and (cheminAccesLu <> '') then
      if peNomValeurRegistre = cstRegR3Serveur then
        result:= cheminAccesLu    // chemin forc�ment complet
      else
        result:= includeTrailingPathDelimiter(cheminAccesLu) + DonneCheminAcces(peNomValeurRegistre,false)
        // appel r�cursif � un seul niveau d'imbrication, avec le mode "chemin incomplet"
    else    // pas normal de ne pas trouver cette valeur
      EXCEPTIONINTERNE(defErr300,format(lbErrCleIntrouv,[cstRegR3Serveur,cstRegCheminsDAcces]))

  else
  begin
    if LitRegistreADN(cstRegCheminsDAcces,peNomValeurRegistre,cheminAccesLu)
     and (cheminAccesLu <> '') then
      result:= cheminAccesLu
    else  // anciennes versions : il n'y a pas ces valeurs de registre, les noms des dossiers sont fig�s
      if peNomValeurRegistre = cstRegPlans then
        result:= cstNomDossierPlansAncien
      else
        if peNomValeurRegistre = cstRegSymboles then
          result:= cstNomDossierSymbolesAncien
        else
          EXCEPTIONINTERNE(defErr158);   // param�tre peNomValeurRegistre pas pr�vu !
  end;

  SortieProc;
end;


{ --------------------------------------------------------------------------------------------- }
procedure clsReqGenerique.ControleVersionDansBD;
{ Contr�le la version enregisr�e dans la table Param�tres avec la partie "principale" de cstVersion
  D�clenche une exception si ce n'est pas conforme }
{ --------------------------------------------------------------------------------------------- }

var versionLue: tabAttrib;

begin
  EntreeProc(ClassName+'.LitVersionDansBD');
  { Contr�le de version (3.4.3) }
  LitParamChaine([defParVersionR3Web],[''],versionLue);
  if versionLue[0] = '' then
    EXCEPTIONEXTERNE (lbErrVersionNonEnregistree+sautDeLigne+lbErrR3WebIncorrectementInstalle);

  if leftStr(cstVersion,length(versionLue[0])) <> versionLue[0] then
    EXCEPTIONEXTERNE(lbErrVersionIncorrecte+sautDeLigne
     +lbErrR3WebIncorrectementInstalle+sautDeLigne
     +format(lbVersionServeurR3Web,[cstVersion])+sautDeLigne
     +format(lbVersionBD,[versionLue[0]]));

  SortieProc;
end;


{ --------------------------------------------------------------------------------------------- }
procedure clsReqGenerique.ControleVersionClient(
 const peVersionClient: string);  // version client re�ue en attribut de requ�te
{ Contr�le compatibilit� versions client et serveur                                             }
{ --------------------------------------------------------------------------------------------- }
var
  libelleVersionClient,versionClient,versionServeur: string;

begin
  EntreeProc(ClassName+'.ControleVersionClient');
    { Enlever toutes les lettres (il peut y en avoir deux) � droite de la version client }
    versionClient:= peVersionClient;
    while (versionClient <> '')
     and (rightStr(versionClient,1) < '0') or (rightStr(versionClient,1) > '9') do
      versionClient:= leftStr(versionClient,pred(length(versionClient)));
    { M�me traitement pour la version serveur }
    versionServeur:= cstVersion;
    while (versionServeur <> '')
     and (rightStr(versionServeur,1) < '0') or (rightStr(versionServeur,1) > '9') do
      versionServeur:= leftStr(versionServeur,pred(length(versionServeur)));

    if versionServeur <> versionClient then
    begin
      if ReqAdmin then
        libelleVersionClient:= lbVersionModuleAdministration
      else
        libelleVersionClient:= lbVersionModuleUtilisateur;
      RAISE excArretTotal.Create(lbErrVersionClientIncompatible+sautDeLigne
       +format(libelleVersionClient,[peVersionClient])+sautDeLigne
       +format(lbVersionServeurR3Web,[cstVersion]));
    end;

  SortieProc
end;

{ --------------------------------------------------------------------------------------------- }
function clsReqGenerique.Diagnostic(peChRequete: string): string;
{ Test du dialogue client-serveur, base de registres et BD }
{ --------------------------------------------------------------------------------------------- }

var
  buf,cleCherchee,mdp,serveur,nomUtil,nomBase,dataSource,cheminAccesServeur: string;
  reg: TRegistry;
  tabNom,tabVal: tabAttrib;
  presente: boolean;
  nbAccesMaj,nbAccesCon,nbMilliers,codeClient,occupation,tailleAffichee: integer;
  categ: TCategorie;
  oExtraction: clsExtraction;
  resultExtraction: string;

begin
  result:= '';

  { Teste la cha�ne lue par RecupereRequete }
  buf := ValChampXml(defBalInfoSession,peChRequete,tabNom,tabVal,presente);
  if not presente then
    EXCEPTIONINTERNE(defErr201,lbDiagMessageClient)
  else
  begin
    Environnement:= ValeurAttribut(defAttEnvLog,tabNom,tabVal);
    { Teste la base de registres }

    reg:= TRegistry.Create;
    reg.RootKey:= HKEY_LOCAL_MACHINE;
    reg.access:= KEY_READ;
    if Environnement = '' then
      ChercheEnvironnement(reg);    // Met � jour Environnement si un seul environnement a �t� d�fini

    cleCherchee:= cstRegR3Web+Environnement;
    if reg.OpenKey(cleCherchee,false) then
    begin
      cleCherchee:= cleCherchee+'\'+cstCleConnexionBD;
      reg.CloseKey;   // Indispensable sinon pb � la lecture suivante
      if reg.OpenKey(cleCherchee,false) then
        with reg do
        begin
          mdp:= DecodeMdpBase(ReadString('Password'));
          if mdp = '' then
            EXCEPTIONINTERNE(defErr300,format(lbDiagValRegistre,['Password',cleCherchee]))
//            result:= format(lbDiagValRegistre,['Password',cleCherchee])
          else
          begin
            serveur:= ReadString('Provider');
            if serveur = '' then
              EXCEPTIONINTERNE(defErr300,format(lbDiagValRegistre,['Provider',cleCherchee]))
 //             result:= format(lbDiagValRegistre,['Provider',cleCherchee])
            else
            begin
              nomUtil:= ReadString('User ID');
              if nomUtil = '' then
                EXCEPTIONINTERNE(defErr300,format(lbDiagValRegistre,['User ID',cleCherchee]))
//                result:= format(lbDiagValRegistre,['User ID',cleCherchee])
              else
              begin
                nomBase:= ReadString('Initial Catalog');
                if nomBase = '' then
                  EXCEPTIONINTERNE(defErr300,format(lbDiagValRegistre,['Initial Catalog']))
//                  result:= format(lbDiagValRegistre,['Initial Catalog'])
                else
                begin
                  dataSource:= ReadString('Data Source');
                  if dataSource = '' then
                    EXCEPTIONINTERNE(defErr300,format(lbDiagValRegistre,['Data Source']))
//                    result:= format(lbDiagValRegistre,['Data Source'])
                  else
                    { connexion � la base de donn�es }
                    try
                      ConnexionBD.ConnectionString := format(
                       'Provider=%s;Persist Security Info=False;User ID=%s;Password=%s;'
                       +'Initial Catalog=%s;Data Source=%s',
                       [serveur,nomUtil,mdp,nomBase,dataSource]);
                      ConnexionBD.LoginPrompt:= false;
                      ConnexionBD.Open;

                    except       // EOleException
                      on e: Exception do
                        RAISE excConnexionBD.Create(lbErrConnexionBD+e.Message)
                    end;

                    LitCleProtec(false,cheminAccesServeur,nbAccesMaj,nbAccesCon,nbMilliers,codeClient);
                    ControleVersionDansBD;    // 3.4.3

//                    if result = '' then
//                    begin
                      ControleTailleBase(nbMilliers*1000,codeClient,occupation,categ);
                      result:= format(lbModuleServeurVersion,[cstVersion]);
                      if nbMilliers < 900 then
                      begin
                        if nbMilliers = 0 then
                          tailleAffichee:= cstTailleBaseMin   // v3.4.8c : taille 0 assimil�e � 100
                        else
                          tailleAffichee:= nbMilliers*1000;

                        result:= result
                         + sautDeLigne + format(lbOccupation,[occupation,tailleAffichee]);
                        if categ <> eCable then
                          result:= result + format(' (%s)',[InfoCateg[categ].Code]);
                      end;
                      { Test export Excel }
                      oExtraction:= clsExtraction.Create(self); // self pour r�cup�rer la connexion � la base de donn�e (ConnexionBD.Open) r�alis�e dans les tests ci-dessus
                      try
                        result:= result + sautDeLigne;
                        resultExtraction:= oExtraction.TestExportExcel;  // g�n�re une exception defErr304 si probl�me avec driver ADO Excel
                        if resultExtraction = '' then
                          result:= result + lbDiagnosticCorrect
                        else
                          result:= result + resultExtraction; 
                      finally
                        oExtraction.Free;
                      end;
                      result:= RemplitMessage(eMessInfo,result);
                end;
              end;
            end;
          end;
        end
      else
        EXCEPTIONINTERNE(defErr399,format(lbDiagCleRegistre,[cleCherchee]));
      reg.CloseKey;
    end
    else
      EXCEPTIONINTERNE(defErr399,format(lbDiagCleRegistre,[cleCherchee]));
  end;
end;

{ --------------------------------------------------------------------------------------------- }
function clsReqGenerique.DonneNomNiveauSite: string;
{ Donne le nom du premier niveau de lieu (non intersite) }
{ --------------------------------------------------------------------------------------------- }
begin
  EntreeProc(className+'.DonneNomNiveauSite');
  result:='';
  with Query do
  begin
    SQL.Text:= format(
     'Select nom from type where Categorie=''%s'' and NumNiveau=1 and Id_type<>%d',
      [InfoCateg[eLieuSimple].Code,cstIdTypeLieuIntersite]);
    Open;
    if Eof then
      EXCEPTIONINTERNE(defErr124,lbErrNiveau1LieuManquant)
    else
      result:= FlbMajusc1(fieldByName('Nom').AsString);
    Close;
  end;
  SortieProc;
end;

{ --------------------------------------------------------------------------------------------- }
function clsReqGenerique.DonneNomNiveauSite
 (var psGenreGr: TGenreGr)
 : string;
{ Variante de la pr�c�dente qui admet aussi un param�tre genre (masculin ou f�minin)            }
{ --------------------------------------------------------------------------------------------- }

var iG: tGenreGr;

begin
  EntreeProc(className+'.DonneNomNiveauSite');
  result:='';
  with Query do
  begin
    SQL.Text:= format(
     'Select Nom,Genre from type where Categorie=''%s'' and NumNiveau=1 and Id_type<>%d',
      [InfoCateg[eLieuSimple].Code,cstIdTypeLieuIntersite]);
    Open;
    if Eof then
      EXCEPTIONINTERNE(defErr124,lbErrNiveau1LieuManquant);
    if fieldByName('Genre').AsString <> '' then    // ce test emp�che de planter en cas d'incoh�rence BD
      for iG:= Low(TGenreGr) to high(TGenreGr) do
        if fieldByName('Genre').AsString[1] = codeGenreGr[iG] then
          psGenreGr:= iG;
    result:= fieldByName('Nom').AsString;
    Close;
  end;
  SortieProc;
end;

{ --------------------------------------------------------------------------------------------- }
function clsReqGenerique.DansLignee(
 peCategorie: TCategorie;   // Cat�gorie de l'�l�ment � rechercher
 peIdent: integer;     // Identifiant de l'�l�ment � rechercher
 peLignee: string)  // lign�e sous forme XML telle que renvoy�e par LigneeItemLieux
 : boolean;  // Renvoie vrai si l'�l�ment (peCategorie,peIdent) est pr�sent dans la lign�e peLignee
{ --------------------------------------------------------------------------------------------- }

begin
  result:= false;
  while peLignee <> '' do
  begin
    peLignee:= ValChampXml(defBalItem,peLignee);
    if (ValChampXml(defBalCategorie,peLignee) = InfoCateg[peCategorie].Code)
     and (StrToInt(ValChampXml(defBalNum,peLignee)) = peIdent) then
    begin
      result:= true;
      BREAK
    end;

    peLignee:= ValChampXml(defBalFils,peLignee);  // modifie une COPIE de peLignee car pas transmis par VAR
  end;
end;

{ ----------------------------------------------------------------------------- }
function clsReqGenerique.DateR3WebClient(
 peDate: TDateTime)   // date telle que lue dans la base
 : string;            // Renvoie une cha�ne JJ/MM/AAAA
{ ----------------------------------------------------------------------------- }
{ NB: peDate peut �tre transmise sous forme d'un entier (date sans la partie heure) }

begin
  EntreeProc(ClassName+'.DateR3WebClient');
  if peDate = 0 then
    result:= ''    // sinon il renverrait '30/12/1899'
  else
    result:= dateToStr(peDate,FormatsADN);
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
function clsReqGenerique.DateR3WebServeur(
 peChaineDate: string)   // cha�ne en format JJ/MM/AAAA ou JJ/MM (telle qu'elle est transmise par le client)
 : integer;     // Renvoie une valeur stockable en base (= partie enti�re du TDateTime correspondant � cette date)
{ ----------------------------------------------------------------------------- }

begin
  EntreeProc(ClassName+'.DateR3WebServeur');
  result:= trunc(strToDateDef(peChaineDate,0,formatsADN));
  SortieProc;
end;

{ ---------------------------------------------------------------------------- }
procedure clsReqGenerique.ChargeXmlToTab(pesoTa: clsTabAsso; peBalise, peAtt, peXml: string);
{ pour une suite de balises identiques charge le tableau associatif pesoTa avec
  comme indice la valeur de l'attribut et comme valeur la valeur entre balise
}
{ ---------------------------------------------------------------------------- }
var
presente : boolean; nomAtt, nomChamp : tabAttrib;
valeurChamp, nom : string;

begin
  EntreeProc(ClassName+'.ChargeXmlToTab');
  pesoTa.Effacer;     // v3.6.0: emp�che la persistance des tableaux de champs quand  plusieurs appels successifs � Maj
  repeat
    valeurChamp:= ValChampXml(peBalise, peXml, nomAtt, nomChamp, presente,1,true);
    if presente then    // v3.5.3 : ne pas utiliser nomAtt et nomChamp si la balise n'est pas trouv�e
    begin
      nom := ValeurAttribut(peAtt,nomAtt,nomChamp);
      if nom<>'' then
        pesoTa[nom] := RendXml(trim(valeurChamp));
    end
    else       // simplification v3.5.3
      BREAK
  until false;
  SortieProc;
end;


{ ---------------------------------------------------------------------------- }
procedure clsReqGenerique.AncetresConnecteur(
 peIdCnc: integer;  // Identifiant de connecteur
 var psNomLt: string;   // nom du local technique (si pas connecteur de terminaison)
 var psNomGe: string;   // nom du groupe d'�quipements s'il existe (si pas connecteur de terminaison)
 var psNomEqt: string); // nom de l'�quipement  ou de la terminaison
{ Renvoie les noms du LT, du GE, de l'�quipement ou de la terminaison d'appartenance
 du connecteur peIdCnc }
{ ---------------------------------------------------------------------------- }

var
  QLignee: TADOQuery;

begin
  EntreeProc(ClassName+'.AncetresConnecteur');
  try
    QLignee:= CreeADOQuery();
    with QLignee do
    begin
//      Connection:= ConnexionBD;
      SQL.Add('select E.Nom as NomEqt, LP.Nom as NomLP, LGP.Nom as NomLGP,');
      SQL.Add('TE.Categorie as CatEqt, TLP.Categorie as CatLP');
      SQL.Add(format('from %s C join %s E on E.Id_eqt = C.Id_eqt',
       [NomTableCompo[eConnecteur],NomTableCompo[eEquipement]]));
      SQL.Add('join Type TE on TE.Id_type = E.Id_type');
      SQL.Add(format('join %s LP on LP.Id_lieu = E.Id_lieu',
       [NomTableCompo[eGroupe]]));
      SQL.Add('join Type TLP on TLP.Id_type = LP.Id_type');
      SQL.Add(format('join %s LGP on LGP.Id_lieu = LP.IdPere',
       [NomTableCompo[eLocalTechnique]]));
      SQL.Add(format('where C.Id_cnc = %d',[peIdCnc]));

      Open;
      if Eof then
        EXCEPTIONINTERNE(defErr19);
      psNomEqt:= fieldByName('NomEqt').AsString;
      if fieldByName('CatEqt').AsString = InfoCateg[eTerminaison].Code then
      begin
        psNomLt:= '';
        psNomGe:= '';
      end
      else
        if fieldByName('CatLP').AsString = InfoCateg[eGroupe].Code then
        begin
          psNomLt:= fieldByName('NomLGP').AsString;
          psNomGe:= fieldByName('NomLP').AsString
        end
        else
        begin
          psNomLt:= fieldByName('NomLP').AsString;
          psNomGe:= '';
        end;
    end;
  finally
    QLignee.Free;
  end;
  SortieProc;
end;

{ ---------------------------------------------------------------------------- }
procedure clsReqGenerique.LitCleProtec(
 peAppelParImport: boolean;   // true si appel en mode Exe (Import ou TestRequetes)
 var psCheminAccesServeur: string;  // chemin d'acc�s au serveur -- peut avoir une valeur initiale dans le contexte d'import
 var psNbAccesMaj: integer;   // nombre maxi d'acc�s simultan�s en mise � jour
 var psNbAccesCon: integer;   // nombre maxi d'acc�s simultan�s en consultation
 var psNbMilliers: integer;   // taille base de donn�es maxi
 var psCodeClient: integer);  // n� licence client
{ Contr�le cl� de protection }
{ ---------------------------------------------------------------------------- }

const
//  cleMiracle = '5520-41330-40345-9886';    // a �t� chang�e en v3.4.0b
  sourceCleMiracle = 42587;   // remplace la cl� miracle �crite en dur
  cstValeurMaxiCodable = 983;

var
  infosFic: TSearchRec;
  chProdId,quatreCles,cheminTemp: string;
  nbSource,posit1,posit2,numProd,cle1,cle2,cle3,cle4,iter: integer;
  resultFind,partieCle,d1,d2,d3,d4: integer;
  chCleMiracle: string;

begin
  EntreeProc(ClassName+'.LitCleProtec');

  if Provider = '' then
  begin
    if not LitRegistreADN('',cstRegLicence,quatreCles) then  // enregistrement cl� conforme 3.4.7a
      EXCEPTIONEXTERNE(format(lbErrEnvLogNonTrouve,[Environnement]));
      // v3.5.2b (1099): Environnement logique non trouv� (en import, maintenance,etc. ce peut �tre une erreur de saisie de l'env. logique)
    { Attention, en 3.5.2b, LitRegistreADN renvoie true si la cl� est trouv�e mais pas la valeur (ou valeur vide)
      donc si on a trouv� la cl� licence mais avec une valeur vide, il faut rechercher la "valeur par d�faut" de la cl� corresp. � l'env.log. }
    if (quatreCles = '') and not LitRegistreADN('','',quatreCles) then  // compatibilit� avec les anciennes versions
      EXCEPTIONEXTERNE(format(lbErrEnvLogNonTrouve,[Environnement]));
      // v3.5.2b (1099): Environnement logique non trouv� (en import, maintenance,etc. ce peut �tre une erreur de saisie de l'env. logique)

     { v3.5.2b (1099):LitRegistreADN ne renvoie false que si la cl� n'existe pas }
    if quatreCles = '' then
      // On a trouv� l'environnement logique mais pas la licence
      EXCEPTIONINTERNE(defErr399,format(lbCodeErreur,[0]));

    if not LitRegistreADN(cstRegCheminsDAcces,cstRegR3Serveur,psCheminAccesServeur)
     or (psCheminAccesServeur = '') then   // v3.5.2b (??1) car LitRegistreADN ne renvoie false que si la cl� n'existe pas
      EXCEPTIONINTERNE(defErr300,lbErrCheminReponse)
  end;

  { C1) Trouver l'heure de cr�ation du dossier de R3Web }
  cheminTemp:= psCheminAccesServeur;
  if (psCheminAccesServeur = '') or (psCheminAccesServeur[length(psCheminAccesServeur)]<>'\') then
    cheminTemp:= psCheminAccesServeur + '\TEMP'
  else
    cheminTemp:= psCheminAccesServeur+'TEMP';
  resultFind:= FindFirst(cheminTemp,faDirectory,infosFic);

  if (resultFind <> 0) and not peAppelParImport then
    EXCEPTIONINTERNE(defErr301,format('(%d)',[resultFind]));
    // NB: on obtient parfois un code 5 = ERROR_ACCESS_DENIED (Atlantica)

  if not peAppelParimport then
  begin
    nbSource:= infosFic.FindData.ftCreationTime.dwLowDateTime div 10 mod 100000;

    // 6 derniers chiffres de la date/heure de cr�ation sauf le dernier (qui est toujours pair donc sans doute pas tr�s significatif)
    SysUtils.FindClose(infosFic);

    { C2) Lire la cl� Product Id et l'ajouter aux secondes }
    numProd:= 123456;
    if LitRegistres(cstRegCurrentVersion,cstRegProductId,chProdId) then
      if chProdId <> '' then     // v3.5.2b (??1)
      begin
        posit1:= pos('-',chProdId);
        posit1:= posEx('-',chProdId,succ(posit1));  // cherche la 2e occurrence
        if posit1 > 0 then
          posit2:= posEx('-',chProdId,succ(posit1));  // cherche la 3e occurrence
          if (posit2 > 0) and (posit2 - posit1 > 6) then
            { prend les 6 chiffres de droite du Product Id}
            numProd:= strToIntDef(copy(chProdId,posit2-6,6),123456)
      end;

    nbSource:= (nbSource+numProd) mod 1000000;    // = nb de 6 chiffres maximum
    posit1:= pos('.',cstVersion);
    posit2:= posEx('.',cstVersion,succ(posit1));
    nbSource:= nbSource + trunc(strToFloat(leftStr(cstVersion,pred(posit2)),FormatsADN)*1000);
    // Ajout du d�but du n� de version (ex 3.3.1 ==> 3300) - ceci laisse la possibilit� de 3 chiffres apr�s le premier point

  end;

  partieCle:= AlgoProtec(sourceCleMiracle);
  chCleMiracle:= intToStr(partieCle);          // premi�re section de la cl�
  for iter:= 2 to 4 do
  begin
    partieCle:= AlgoProtec(partieCle);
    chCleMiracle:= chCleMiracle + '-' + intToStr(partieCle);
  end;
  // Valeur finale doit �tre = '189534-143743-19055-184839'

  { C3: Contr�ler la coh�rence de la valeur de la cl� correspondant � l'environnement logique }
  if peAppelParImport or (quatreCles = chCleMiracle) then
    { NB: import v3.3.1b: plus de contr�les: les var. de contr�le sont mises � leur maximum }
  begin
    psNbAccesMaj:= cstValeurMaxiCodableDansLicence;
    psNbAccesCon:= cstValeurMaxiCodableDansLicence;
    psNbMilliers:= cstValeurMaxiCodableDansLicence;
    psCodeClient:= 0;
  end
  else
  begin

    posit1:= pos('-',quatreCles);
    if posit1 <= 0 then
      EXCEPTIONINTERNE(defErr399,format(lbCodeErreur,[1]));

    if not tryStrToInt(leftStr(quatreCles,pred(posit1)),cle1) then
      EXCEPTIONINTERNE(defErr399,format(lbCodeErreur,[2]));

    posit2:= posEx('-',quatreCles,succ(posit1));
    if posit2 <= succ(posit1) then
      EXCEPTIONINTERNE(defErr399,format(lbCodeErreur,[4]));
    if not tryStrToInt(midStr(quatreCles,succ(posit1),pred(posit2-posit1)),cle2) then
      EXCEPTIONINTERNE(defErr399,format(lbCodeErreur,[5]));

    posit1:= posit2;
    posit2:= posEx('-',quatreCles,succ(posit1));
    if posit2 <= succ(posit1) then
      EXCEPTIONINTERNE(defErr399,format(lbCodeErreur,[7]));
    if not tryStrToInt(midStr(quatreCles,succ(posit1),pred(posit2-posit1)),cle3) then
      EXCEPTIONINTERNE(defErr399,format(lbCodeErreur,[8]));

    if not tryStrToInt(rightStr(quatreCles,length(quatreCles)-posit2),cle4) then
      EXCEPTIONINTERNE(defErr399,format(lbCodeErreur,[10]));

    { 3.5.6 - on a ajout� aux cl�s un 1er chiffre redondant issu des 4 cl�s modulo 200000
      on isole les valeurs "significatives" (= cl�s modulo 200000) des codes de contr�le (= cl�s div 200000)}
    { valeurs redondantes de contr�le }
    d1:= cle1 div 200000;
    d2:= cle2 div 200000;
    d3:= cle3 div 200000;
    d4:= cle4 div 200000;

    { valeurs significatives des cl�s (�quivalent versions < 3.6.5)}
    cle1:= cle1 mod 200000;
    cle2:= cle2 mod 200000;
    cle3:= cle3 mod 200000;
    cle4:= cle4 mod 200000;

    if ((d4*5+d3)*5+d2)*5+d1 <> (cle1+cle2+cle3+cle4) mod 625 then    // 625 = 5*5*5*5
      EXCEPTIONINTERNE(defErr399,format(lbCodeErreur,[12]));
    // [d4d3d2d1] aurait d� �tre l'�quivalent en base 5 de (cle1+cle2+cle3+cle4) mod (5 puissance 4)

   { v3.6.5 - les calculs des caract�ristiques licence sont regroup�s apr�s avoir isol� toutes les cl�s, c'est plus clair }
    psNbAccesMaj:= cle1 - AlgoProtec(nbSource);
    if (psNbAccesMaj <= 0) or (psNbAccesMaj > cstValeurMaxiCodableDansLicence) then
      EXCEPTIONINTERNE(defErr399,format(lbCodeErreur,[3]));

    psNbAccesCon:= cle2 - AlgoProtec(cle1);
    if (psNbAccesCon < 0) or (psNbAccesCon > cstValeurMaxiCodableDansLicence) then
      EXCEPTIONINTERNE(defErr399,format(lbCodeErreur,[6]));

    psNbMilliers:= cle3 - algoProtec(cle2);
    if (psNbMilliers < 0) or (psNbMilliers > cstValeurMaxiCodableDansLicence) then  // modif 3.4.8c: on autorise le cas psNbMilliers = 0 (versions d'essai ou d'�tude)
      EXCEPTIONINTERNE(defErr399,format(lbCodeErreur,[9]));

    psCodeClient:= cle4 - algoProtec(cle3);
    if (psCodeClient <= 0) or (psCodeClient > cstValeurMaxiCodableDansLicence) then
      EXCEPTIONINTERNE(defErr399,format(lbCodeErreur,[11]));

    { Test de la langue en fonction de la licence }
(* Supprim� en v3.6.5
   if lbCodeLangue = 'FR' then
    begin
      { Ici ajouter les contr�les pour le clients qui n'ont pas droit au fran�ais }
      if psCodeClient = 98765 then
        EXCEPTIONEXTERNE(lbErrVotreLicenceNestPasValidePourLaVersionFrancaise)
    end
    else
      if lbCodeLangue = 'EN' then
        if not (psCodeClient in [207,227,229]) then
          // Bureau international du travail ou Eurotunnel ou Yemen LNG
          EXCEPTIONEXTERNE(lbErrYourLicenseIsNotValidForEnglishVersion);   // libell� non traduit !
*)
  end;
  SortieProc;
end;

{ ---------------------------------------------------------------------------- }
{ Contr�le taille base de donn�es autoris�e:
  Lanc� lors de la cr�ation d'une nouvelle session et du r�veil d'une session  }
function clsReqGenerique.ControleTailleBase(
 peTailleBase: integer;     // taille autoris�e
 peCodeClient: integer;   // code client (pour shunter certains contr�les)
 var psOccupation: integer;  // nb de (pseudo-)connecteurs ou de fonctions selon ce qui approche le plus de peTailleBase
 var psCategCause: TCategorie)   // indicateur cause de d�passement : T si terminaisons, F si fonctions
 : boolean;   // true si la taille autoris�e est d�pass�e
{ ---------------------------------------------------------------------------- }

var
  reqTypesTer1,reqTypesTer2,reqComptage: TADOQuery;
  partieCommune: TStringList;

begin
  EntreeProc(ClassName+'.ControleTailleBase');
  { v3.4.8c: tailleBase � 0 est assimil�e � 100 }
  if peTailleBase =  0 then
    peTailleBase:= cstTailleBaseMin;

  if peTailleBase >= 900000 then  // taille illimit�e: on gagne du temps... (v3.6.3)
    result:= false
  else
  begin
    reqTypesTer1:= CreeADOQuery();
    reqTypesTer2:= CreeADOQuery();
    reqComptage:= CreeADOQuery();
    partieCommune:= TStringList.Create;
    try
      { 1) reqTypesTer1 calcule le "poids partiel" de chaque type de terminaison
       pour les connecteurs de 1 paire ou 1 fil }
      with reqTypesTer1 do
      begin
        SQL.Add('select (count(C.Id_cnc) -1)/4 +1 as compteCnc,E.Id_type from Equipement E');
        // pour 1,2 ou 4 connecteurs d'une paire ou d'un fil �a compte toujours 1
        partieCommune.Add('join Connecteur C on C.Id_eqt=E.Id_eqt');
        partieCommune.Add('join TypeCnc TC on TC.Id_typC=C.Id_typC');
        partieCommune.Add('join type T on T.Id_Type=E.Id_type');
        partieCommune.Add(format('where E.Id_lieu is null and E.Id_Type is not null and T.Categorie = ''%s''',
         [InfoCateg[eTerminaison].Code]));
        SQL.AddStrings(partieCommune);
        SQL.Add('and TC.NbGroupes = 1 and TC.TailleGroupe <= 2');
        SQL.Add('group by E.Id_eqt,E.Id_type');
        // NB: les noms de tables sont fixes car ce ne sont pas les copies temporaires pour ODT
        // et le tableau NomTableCompo n'est pas initialis� dans le cas o� l'appelant est Diagnostic
        Open;
      end;

      { 2) reqTypesTer2 calcule le "poids partiel" de chaque type de terminaison
       pour les connecteurs de plusieurs groupes ou plus d'une paire par groupe }
      with reqTypesTer2 do
      begin
        SQL.Add('select count(C.Id_cnc) as CompteCnc,E.Id_type from Equipement E');
        SQL.AddStrings(partieCommune);
        SQL.Add('and (TC.NbGroupes > 1 or TC.TailleGroupe > 2)');
        SQL.Add('group by E.Id_eqt,E.Id_type');
        Open;
      end;

      { 3) Nombre de terminaisons vraies existantes pour chaque type }
      with reqComptage do
      begin
        SQL.Add('select count(E.Id_type) as CompteTer,E.Id_type from Equipement E');
        SQL.Add('join Type T on T.Id_type=E.Id_type');
        SQL.Add(format('where T.Categorie = ''%s'' and E.Id_lieu is not null',
         [InfoCateg[eTerminaison].Code]));
        // NB: inutile de tester Id_type is not null (pour exclure les pr�cr�ations ODT) gr�ce � la jointure sur Types
        { v3.6.4b (1288 - 1289) }
        SQL.Add('and not exists(');
        SQL.Add('select Id_eqt from Connecteur Cn');
        SQL.Add('join Cablage C1 on C1.Id_cnc = Cn.Id_cnc');
        SQL.Add('join Cablage C2 on C2.Id_cnc = Cn.Id_cnc and C2.Cnx = C1.Cnx');    // autre lien de c�blage sur la m�me connexion
        SQL.Add('where Cn.Id_eqt = E.Id_eqt and C2.Id_lien <> C1.Id_lien)');
        SQL.Add('group by E.Id_type');
        Open;

        { 4) Comptage du poids total des terminaisons existantes }
        psOccupation:= 0;
        while not Eof do
        begin
          { On fait les 2 recherches � la fois car un type de terminaison peut contenir
            � la fois des connecteurs d'une paire ou d'un fil et des connecteurs de structure diff�rente }
          if reqTypesTer1.Locate('Id_type',FieldByName('Id_type').asInteger,[]) then
            psOccupation:= psOccupation +
             reqTypesTer1.FieldByName('CompteCnc').asInteger * FieldByName('CompteTer').AsInteger;
          if reqTypesTer2.Locate('Id_type',FieldByName('Id_type').asInteger,[]) then
            psOccupation:= psOccupation +
             reqTypesTer2.FieldByName('CompteCnc').AsInteger * FieldByName('CompteTer').AsInteger;
          Next
        end;
      end;

      psCategCause:= eTerminaison;    // crit�re a priori = terminaisons
      if psOccupation > peTailleBase then
        result:= true
      else
      begin
        { Comptage des fonctions }
        with reqComptage do
        begin
          SQL.Clear;
          SQL.Add('select count(IdSite) as Compte from Fonction');
          // La requ�te exclut les lignes avec IdSite � null: mod�les ou fonctions pr�cr��es par ODT
          Open;
          if FieldByName('Compte').AsInteger > psOccupation then
          begin
            psOccupation:= FieldByName('Compte').AsInteger;
            psCategCause:= eFonction;
          end;
        end;

        if psOccupation > peTailleBase then
          result:= true
        else
          { Comptage des c�bles (sauf pour clients 42 (ARCELOR FOS), 150 (CESNAC) 186 (LMJ)) }
          if (peCodeClient = 150) or (peCodeClient = 42) or (peCodeClient = 186) then
            result:= false
          else
          begin
            with reqComptage do
            begin
              SQL.Clear;
              SQL.Add('select count(IdSite) as Compte from Cable');
              // La requ�te exclut les lignes avec IdSite � null: mod�les ou fonctions pr�cr��es par ODT
              Open;
              if FieldByName('Compte').AsInteger*0.8 > psOccupation then
              begin
                psCategCause:= eCable;
                psOccupation:= trunc(FieldByName('Compte').AsInteger*0.8);
              end;
            end;
            result:=  psOccupation > peTailleBase;
          end;
      end;

    finally
      reqTypesTer1.Free;
      reqTypesTer2.Free;
      reqComptage.Free;
      partieCommune.Free;
    end;
  end;

  SortieProc;
end;

{ ---------------------------------------------------------------------------- }
{ Contr�le nombre d'acc�s simultan�s autoris�s }
procedure clsReqGenerique.ControleNbAcces(
 var pesProfilModif: boolean;  // vrai si profil de la session est en modification
     // peut �tre mis � false si le nombre d'acc�s en mise � jour est atteint
 peNbAccesMaj,peNbAccesCon: integer;   // nombres d'acc�s autoris�s en mise � jour et en consultation
 peTopHorloge: TDateTime;   // top horloge actuel
 peDelaiVeille,peDureeMaxRequete: integer);   // valeurs des param�tres g�n�raux d�lai veille et dur�e max requ�te
{ ---------------------------------------------------------------------------- }

var
  compteAccesMaj: integer;

begin
  EntreeProc(ClassName+'.ControleNbAcces');
  if peNbAccesMaj < cstValeurMaxiCodableDansLicence then    // si nombre d'acc�s limit�
    with Query do
    begin
      SQL.Clear;
      SQL.Add('select * from Session');
      SQL.Add('where ' +CritereSessionActive(peTopHorloge,peDelaiVeille,peDureeMaxRequete));
      SQL.Add(format('and not Preferences in (''%s'',''%s'')',[cstProfilModuleAdmin,cstProfilTelMobile]));
      // On exclut les sessions du module administrateur (v3.6.1: et la consultation depuis t�l�phone mobile)
      Open;
      if RecordCount >= peNbAccesCon + peNbAccesMaj then
        EXCEPTIONEXTERNE(lbErrNbMaxSessionsAtteint);
      compteAccesMaj:= 0;
      if pesProfilModif then
      begin
        while not Eof do
        begin
          if not (FieldByName('Preferences').AsString = '')  // cas en principe impossible sauf pb de maj de version
           or (FieldByName('Preferences').AsString[1] in [cstProfilModif,cstProfilCreationInhibee]) then
            inc(compteAccesMaj);
          Next
        end;
        if compteAccesMaj >= peNbAccesMaj then  // si trop de sessions en mise � jour
          { passage temporaire en consultation }
          pesProfilModif:= false;
      end;
    end;
  SortieProc
end;


{ ---------------------------------------------------------------------------- }
function clsReqGenerique.LitCheminImportExport(
 peNumeroParametre: integer)     //  =  defParCheminAccesImport ou defParCheminAccesExport
 : string;
{ Lit le chemin o� l'on doit cr�er les fichier d'export et d'import }
{ ---------------------------------------------------------------------------- }
var
  tabParGene: tabAttrib;
  nomDossier,cheminAcces,partieChemin: string;

begin
  EntreeProc(ClassName+'.LitCheminImportExport');
  result:= '';

  { Lecture chemin d'acc�s au serveur }
  cheminAcces:= includeTrailingPathDelimiter(DonneCheminAcces(cstRegR3Serveur));

  { Lecture param�tre g�n�ral "Chemin d'acc�s de l'export" (n�11) ou "Chemin d'acc�s de l'export" (n�18)"}
  LitParamChaine([peNumeroParametre],[''],tabParGene);
  { v3.5.6: on n'�crit plus que des chemins relatifs dans les param�tres chemin d'acc�s import et export
   (le chemin doit �tre un sous-dossier du chemin d'acc�s au serveur) }
  if (tabParGene[0] = '') or (tabParGene[0] = '\') then   // param�tre absent ou mal rempli
  begin
    case peNumeroParametre of
      defParcheminAccesImport:
        nomDossier:= cstNomDossierImportParDefaut;
      defParCheminAccesExport:
        nomDossier:= cstNomDossierExportParDefaut
      else
        EXCEPTIONINTERNE(DefErr153);
    end;
    result:= cheminAcces + nomDossier + '\';    // 3.5.6b
  end
  else
  begin
    nomDossier:= excludeTrailingPathDelimiter(tabParGene[0]);
    // le cas tabParGene[0] = '\' est trait� plus haut donc nomDossier ne peut �tre mis � ''

(*
    { v3.5.6a: Dans les anciennes versions, ils avaient le droit de mettre n'importe quel chemin et pas juste un nom de dossier }
    partieChemin:= extractFilePath(nomDossier);
    if partieChemin <> '' then
    begin
      { Si le chemin des exports (resp. imports) inclut le chemin du serveur }
      if leftStr(uppercase(partieChemin),length(cheminAcces)) = uppercase(cheminAcces) then
        // Ce chemin est celui du serveur
        nomDossier:= extractFileName(nomDossier)    // on fait comme si c'�tait un nom de fichier, pour pouvoir se servir de ExtractFileName
      else
      begin
        case peNumeroParametre of
          defParcheminAccesImport:
            libelleComplementaire:= lbDImport;
          defParCheminAccesExport:
            libelleComplementaire:= lbDExport;
        end;
        EXCEPTIONEXTERNE(format(lbErrChangerCheminImportOuExport,[libelleComplementaire]));
      end;
    end;   *)

    { 3.5.6b - Le chemin n'est qu'un sous-dossier et ne doit pas �tre un chemin complet (l'install le corrige) }
    result:= cheminAcces + nomDossier + '\';
  end;

  SortieProc;
end;


{ ---------------------------------------------------------------------------- }
function clsReqGenerique.ListeLieuxNiveau1: string;
{ Donne la liste de tous les lieux de niveau 1 ("sites") }
{ ---------------------------------------------------------------------------- }

begin
  EntreeProc(ClassName+'.ListeLieuxNiveau1');
  with Query do
  begin
    SQL.Clear;
    SQL.Add(format('select L.Nom,L.Id_lieu from %s L join Type T on T.Id_type = L.Id_type',
     [NomTableCompo[eLieuSimple]]));
    if not Administrateur then
      SQL.Add(format('left join Droit_Lieu D on D.Id_prof=%d and D.Id_lieu=L.Id_lieu',
       [IdProfil]));   // 3.6.3a (1274)
    SQL.Add(format('where L.IdPere is not null and T.Categorie = ''%s'' and T.NumNiveau = 1',
     [InfoCateg[eLieuSimple].Code]));   // modifi� 3.3.4b : exclut les mod�les
//    SQL.Add('order by L.Classement');  // modif 3.3.4c
    if ValeurParam(defAttAvecInterSite) <> 'O' then
      SQL.Add(format('and T.Id_type <> %d',[cstIdTypeLieuIntersite]));
    if not Administrateur then  // 3.6.3a (1274) ignore les sites avec aucun acc�s pour ce profil
      SQL.Add(format('and (D.DroitLieu is null or D.DroitLieu <> %d)',[cstDroitAucunAcces]));
    SQL.Add('order by L.Nom');   // 3.4.3b: classement par nom
    Open;
    result:= '';
    while not Eof do
    begin
      result:= result + RemplitBalise(defBalSite,FieldByName('Nom').AsString,
       [defAttNumero],[fieldByName('Id_lieu').AsString],true);
      Next;
    end;
    result:= RemplitMessage(eMessListe,result);
  end;
  SortieProc;
end;

{ ---------------------------------------------------------------------------- }
{ cr�ation d'une requ�te avec gestion du temps maxi de r�ponse }
function clsReqGenerique.CreeADOQuery(
 peDelaiInfini: boolean = true)  // true s'il faut donner comme timeOut la valeur lue dans la table des param�trage
 : TADOQuery;          // false si la requ�te peut avoir un temps d'ex�cution illimit�
{ ---------------------------------------------------------------------------- }

begin
  EntreeProc(ClassName+'.CreeADOQuery');
  result:= TADOQuery.Create(nil);
  result.Connection:= ConnexionBD;
  result.ParamCheck:= false;     // indispensable pour emp�cher bug n� 1015
  if peDelaiInfini then
    result.CommandTimeOut:= 0
  else
    result.CommandTimeout:= DureeMaxiRequete;    // ne marche pas : ce n'est pris en compte que si c'est � 0
  SortieProc;
end;

{ ----------------------------------------------------------------------------- }
function clsReqGenerique.NomComposantUnique:
 string;  // Renvoie une cha�ne suppos�e unique � partir du num�ro de session et de l'heure
 { ---------------------------------------------------------------------------- }
var topHorloge: tDateTime;

begin
  topHorloge:= Date + getTime;
  result:= intToStr(NumSession)+'#'+FloatToStr(topHorloge);   // cela garantit en pratique (mais pas en th�orie) l'unicit� du nom
end;

{ ----------------------------------------------------------------------------- }
{ Suppression de toutes les op�rations li�es � l'action de num�ro NumAction
{ Contexte: ex�cution ODT - NumAction repr�sente une action � l'�tat 'Pr�vue' qui sera recr��e par l'ex�cution
  ou bien appel par clsOperation.SupprAction : suppression d'une action demand�e par l'utilisateur }
{ Proc�dure remani�e en 3.4.1 }
procedure clsReqGenerique.SupprimeActionPrevue
 (peActionPrincipale: TOperationR3Web;  // code de l'action principale
  peCategorie: TCategorie = eLieuSimple;  // cat�gorie d'objet
  peContexteSupprManu: boolean = false);  // true si l'appel vient de clsOperation.SupprActionOdt (v3.5.3a)
  // si peActionPrincipale = eCreerComposant ou eModifierComposant ou eSupprimerComposant
{ ----------------------------------------------------------------------------- }

{ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ }
{ Effacement des �l�ments pr�cr��s (= ceux qui ont des identifiants) dans les tables r�elles                      }
procedure SQLEffaceEltsPrecrees(
 peCategCompoEff: TCategorie;   // Cat�gorie de l'�l�ment � effacer
 peCestLActionPrincipale: boolean;  // true s'il s'agit de l'action principale � supprimer
 peScriptSQL: TStrings);    // scriptSQL � enrichir

var
  finReq: TStringList;

begin
  finReq:= TStringList.Create;
  with peScriptSQL do
    try
      { Supprimer la ligne correspondant au composant pr�cr�� }
      Add(format('delete %0:s from %0:s C join %1:s HC on HC.%2:s = C.%2:s',
       [NomTableReelle(peCategCompoEff),NomTableHisto(peCategCompoEff),
       NomChampIdent(peCategCompoEff)]));
      if peCestLActionPrincipale then
        finReq.Add(format('where HC.Id_act = %d',[NumAction]))
      else
      begin
        finReq.Add('join Action A on A.Id_act = HC.Id_act');
        finReq.Add(format('where A.ActionContexte = %d',[NumAction]));
      end;
      AddStrings(finReq);

      case peCategCompoEff of
      eEquipement:
        begin
          { Supprimer les lignes correspondant aux connecteurs pr�cr��s }
          Add(format('delete %0:s from %0:s C join %1:s HC on HC.Id_cnc = C.Id_cnc',
           [NomTableReelle(eConnecteur),NomTableHisto(eConnecteur)]));
          AddStrings(finReq);
        end;

      ePlanLieu:
        begin      // v3.6.7 (1341) cas sp�cial d'appel avec ePlanLieu: on efface en plus les trajets pr�cr��s
          { En plus des plans, supprimer les lignes correspondant aux trajets pr�cr��s }
          Add('delete Trajet from Trajet T join HistoTrajet HC on HC.Id_traj = T.Id_traj');
          AddStrings(finReq);
        end;
      end;
    finally
      finReq.Free;

  end;
end;

{ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ }

var
  QMaj: TADOQuery;

begin
  EntreeProc(ClassName+'.SupprimeActionPrevue');
  QMaj:= CreeADOQuery();
  with QMaj do
    try

      { Supprimer les actions de c�blage/d�c�blage/placement/retrait induites (elles seront recr��es) }
      { On ne peut pas mettre des contraintes triggers qui ex�cutent ces suppressions en cascade }
      SQL.Add('delete from HistoDepart where Id_act in (select Id_act from Action');
      SQL.Add(format('where ActionContexte = %d)',[NumAction]));

      SQL.Add('delete from HistoCablage where Id_act in (select Id_act from Action');
      SQL.Add(format('where ActionContexte = %d)',[NumAction]));

      SQL.Add('delete from HistoLien where Id_act in (select Id_act from Action');
      SQL.Add(format('where ActionContexte = %d)',[NumAction]));

      SQL.Add('delete from DetailAction where Id_act in (select Id_act from Action');
      SQL.Add(format('where ActionContexte = %d)',[NumAction]));

      SQL.Add(format('delete from Action where ActionContexte = %d',[NumAction]));

     { Supprimer l'action (si on est en train d'ex�cuter un Odt, une nouvelle va la remplacer) }
      if peActionPrincipale in [eRetirerFonction,eDecabler,eCabler,eAcheminer,ePlacerFonction,eModifierPropBrassage] then
      begin
        SQL.Add(format('delete from HistoDepart where Id_act = %d',[NumAction]));

        SQL.Add(format('delete from HistoCablage where Id_act = %d',[NumAction]));

        SQL.Add(format('delete from HistoLien where Id_act = %d',[NumAction]));

        SQL.Add(format('delete from DetailAction where Id_act = %d',[NumAction]));
      end;

      { 3.5.3a: il existe maintenant des actions de cr�ation "induites" par une action symbolique (eDupliquer)
        il faut supprimer les �l�ments pr�cr��s }
      if peContexteSupprManu and (peActionPrincipale = eDupliquer) then
      begin
        { Supprimer la ligne correspondant au c�ble pr�cr�� }
        SQLEffaceEltsPrecrees(eCable,false,SQL);    // supprime la ligne de Cable correspondant au c�ble pr�cr��
        { Supprimer les lignes correspondant aux �quipements pr�cr��s }
        SQLEffaceEltsPrecrees(eEquipement,false,SQL);  // supprime les lignes de Equipement et Connecteur correspondant aux �quipements pr�cr��s
      end;

      { 3.5.3a (1135) Si contexte = "suppression manuelle d'une action de cr�ation",
        il faut supprimer les �l�ments pr�cr��s }
      if peContexteSupprManu then
      begin               // modif v3.6.7 (1341)
        if peActionPrincipale = eCreerComposant then
          SQLEffaceEltsPrecrees(peCategorie,true,SQL);
        if (peActionPrincipale in [eCreerComposant,eModifierComposant])
         and (peCategorie in [eLieuSimple,eLocalTechnique,eGroupe]) then
          SQLEffaceEltsPrecrees(ePlanLieu,true,SQL);  // cela va supprimer les Trajets et les Plans pr�cr��s
      end;

      if peActionPrincipale in [eCreerComposant,eModifierComposant,eSupprimerComposant] then
      begin
        SQL.Add(format('delete from %s where Id_act = %d',[NomTableHisto(peCategorie),NumAction]));
      end;


      if peActionPrincipale in [eCreerComposant,eModifierComposant] then
      begin  // La suppression de composant ne g�n�re pas d'enreg dans les tables Histo annexes
        if peCategorie = eEquipement then
          SQL.Add(format('delete from HistoConnecteur where Id_act = %d',[NumAction]));  // si c'est un �quipement

        if peCategorie in [eEquipement,eGroupe,eLocalTechnique,eLieuSimple] then
          SQL.Add(format('delete from HistoSymbole where Id_act = %0:d or Id_ActSuppr = %0:d',[NumAction]));  // si c'est un lieu ou �qt
          // 3.6.7 (1341) condition sur Id_ActSuppr ajout�e

        if peCategorie in [eGroupe,eLocalTechnique,eLieuSimple] then
        begin
          { v3.6.0: Nouvelles tables concern�es }
          SQL.Add('delete HET from HistoExtTrajet HET join HistoTrajet HT on HT.Id_traj = HET.Id_traj');
          SQL.Add(format('where HT.Id_act = %d',[NumAction]));
          SQL.Add(format('delete from HistoTrajet where Id_act = %0:d',[NumAction]));
          SQL.Add(format('delete from HistoTrajet_Famille where Id_act = %d or Id_actSuppr = %0:d',[NumAction]));
          // 3.6.7 (1341) condition sur Id_ActSuppr ajout�e
          SQL.Add(format('delete from HistoPlanLieu where Id_act = %d or Id_actSuppr = %0:d',[NumAction]));
          // 3.6.7 (1341) suppression dans HistoPlanLieu carr�ment oubli�e !
        end;
      end;

      if peActionPrincipale in [eCreerComposant,eModifierComposant,eCabler,eModifierPropBrassage] then   // v3.5.4
        SQL.Add(format('delete from HistoValeurProp where Id_act = %d',[NumAction]));


      SQL.Add(format('delete from Action where Id_act = %d',[NumAction]));
      ExecSQL;

    finally
      Free;
    end;
  SortieProc
end;


{ ------------------------------------------------------------------------------------------ }
function clsReqGenerique.DroitModifOdt(   // Donne le  droit de modification sur un ODT
 peProfilOdt: integer)  //  Profil de l'ODT
 : boolean;   // Renvoie true si l'ODT est modifiable, false sinon
{ La fonction utilise les propri�t�s DroitsGeneraux et IdProfil de l'objet clsReqGenerique }
{ ------------------------------------------------------------------------------------------ }

begin
  result:= (DroitsGeneraux and defDOdtMajTousProf <> 0)
   or (DroitsGeneraux and defDOdtMajMonProf <> 0) and (peProfilOdt  = IdProfil)
  // J'ai le droit de Maj de tous les profils
  // ou j'ai le droit de maj des ODT de mon profil et l'ODT poss�de mon profil
end;


{ ------------------------------------------------------------------------------------------ }
{ Mise � jour du taux d'avancement final d'une phase de traitement
  (pour une requ�te utilisant AjouteAvancement)                                               }
procedure clsReqGenerique.MajTauxAvancementMaxi(
 peValeurTaux: TPourcentageEntier);
{ ------------------------------------------------------------------------------------------ }
begin
  if peValeurTaux < TauxAvancement then
    TxAvMax:= TauxAvancement   // cas pas normal
  else
    TxAvMax:= peValeurTaux;
end;

{ ------------------------------------------------------------------------------------------ }
function clsReqGenerique.LitTauxAvancementMaxi: TPourcentageEntier;
{ ------------------------------------------------------------------------------------------ }
begin
  result:= TxAvMax;
end;

{ ------------------------------------------------------------------------------------------ }
{ Mise � jour de l'enreg li� � une requ�te longue donnant des infos sur son avancement (v3.5.6) }
procedure clsReqGenerique.MajAvancement(
 pePourcentage: TPourcentageEntier;
 peLibelle: string = '';        // libell� � mettre dans l'enreg d'avancement
 peMajTxAvMin: boolean = true);    // false s'il ne faut pas mettre � jour TxAvMin
{ ------------------------------------------------------------------------------------------ }

begin
  EntreeProc(ClassName+'.MajAvancement');

  TauxAvancement:= pePourcentage;   // utilis� par AjouteAvancement

  if peMajTxAvMin then
    TxAvMin:= pePourcentage;   // utilis� par AjouteAvancement

  if TauxAvancementMaxi <= TauxAvancement then
    TauxAvancementMaxi:= 100;    // on a n�glig� de remettre � jour le taux d'avancement maxi de la phase en cours

{$IFDEF MODEEXE}
  if @ProcRafraichProgression <> nil then
    ProcRafraichProgression(NumSession);   // contexte MaintenanceR3Web uniquement
{$ENDIF}
  DerniereMajAvancement:= getTime;
  with Command do
  begin
    CommandText:= format(
     'update AvancementRequete set Pourcentage = %d',
      [pePourcentage]);
    if peLibelle <> '' then
      CommandText:= CommandText + format(',Libelle = %s',
       [quotedStr(peLibelle)]);
    // sinon pas de changement du libell� cens� �tre d�j� en place
    CommandText:= CommandText + format(' where Id_ses = %d and Id_req = %d',
     [NumSession,IdRequete]);
    Execute;
  end;
  SortieProc;
end;

{ ------------------------------------------------------------------------------------------ }
{ Ajoute un certain avancement calcul� � partir de l'avancement actuel (v3.6.0)
  et d'un avancement final suppos� d�j� fix� (sinon par d�faut � 100)                        }
procedure clsReqGenerique.AjouteAvancement(
 peProportionAvancementPartiel: real);     // proportion d'avancement pour la phase actuelle seulement
{ ------------------------------------------------------------------------------------------ }

begin
  EntreeProc(ClassName+'.AjouteAvancement');
  if getTime - DerniereMajAvancement > cstIntervalleAvancement then
  begin
    if peProportionAvancementPartiel < 0 then
      peProportionAvancementPartiel:= 0
    else
      if peProportionAvancementPartiel > 1 then
        peProportionAvancementPartiel:= 1;
    MajAvancement(
     TxAvMin + round((TauxAvancementMaxi-TxAvMin) * peProportionAvancementPartiel),
     '',false);   // false <=> ne pas mettre � jour TauxAvDebutPhase
  end;
  SortieProc;
end;


{ ------------------------------------------------------------------------------------------ }
{$IFDEF MODECHRONO}
procedure clsReqGenerique.Mouchard(peTexte: string);
{ ------------------------------------------------------------------------------------------ }

begin
  with TStringList.Create do
  begin
    if FileExists(fichierMouchard) then
      LoadFromFile(fichierMouchard);
    Add(FloatToStr(GetTime()*86400)+' : '+peTexte);
    SaveToFile(fichierMouchard);
    Free;
  end;
end;
{$ENDIF}

end.
