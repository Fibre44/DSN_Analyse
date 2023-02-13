import formidable from "formidable";
import type { NextApiRequest, NextApiResponse } from 'next'
import path from "path";
//import fs from "fs/promises";
import fs from 'fs';
import { DsnParser } from "@fibre44/dsn-parser";
import Excel from 'exceljs'
import { BaseObject, atObject, MobilityObject, ContributionFundObject, ContributionObject, DsnObject, EmployeeObject, EstablishmentObject, MutualEmployeeObject, MutualObject, WorkContractObject, WorkStoppingObject } from "@fibre44/dsn-parser/lib/dsn";
type Data = {
  name?: string,
  error?: string | unknown,
  succes?: string
} | Buffer

type DataDsn = {
  dsn: DsnObject,
  establishment: EstablishmentObject[],
  mutual: MutualObject[],
  contributionFund: ContributionFundObject[],
  employee: EmployeeObject[],
  employeeMutual: MutualEmployeeObject[],
  workContract: WorkContractObject[],
  base: BaseObject[],
  contribution: ContributionObject[]
  rateAt: atObject[],
  rateMobility: MobilityObject[],
  //workStopping: WorkStoppingObject[]
}
export const config = {
  api: {
    bodyParser: false,
  },
};

const saveFile = async (
  req: NextApiRequest,
  date: string,
  saveLocally?: boolean,
): Promise<{ fields: formidable.Fields; files: formidable.Files }> => {
  console.log('save')
  const options: formidable.Options = {};
  if (saveLocally) {
    options.uploadDir = path.join(process.cwd(), `/tmp/${date}`);
    options.filename = (name, ext, path, form) => {
      return Date.now().toString() + "_" + path.originalFilename;
    };
  }
  options.maxFileSize = 4000 * 1024 * 1024;
  options.encoding = 'utf-8'
  const form = formidable(options);
  return new Promise((resolve, reject) => {
    form.parse(req, (err, fields, files) => {
      if (err) reject(err);
      resolve({ fields, files });
    });
  });
};

const removeDir = (date: string): void => {

  fs.rm(process.cwd() + `/tmp/${date}/`, { recursive: true }, (err) => {
    if (err) {
      // File deletion failed
      throw ('Errreur suppression des données')
      ;
    }
    console.log(`Suppression du dossier /tmp/${date}`);

  })

}

export default async function handler(
  req: NextApiRequest,
  res: NextApiResponse<Data>
) {
  try {

    //Etape 1 on test la méthode
    if (req.method != 'POST') {
      return res.status(400).json({ error: 'Vous devez utiliser la méthode POST' })
    }
    //Création du dossier tmp
    if (!fs.existsSync(process.cwd() + "/tmp/")) {
      fs.mkdirSync(process.cwd() + "/tmp/");
    }

    //Etape 2 on va créer un dossier
    const date = Date.now()
    const patch = path.join(process.cwd() + "/tmp/", date.toString())
    const pathString = patch.toString()
    fs.mkdirSync(pathString);
    //Etape 3 on va sauvegarder les fichiers
    await saveFile(req, date.toString(), true,);
    //Etape 4 on liste les fichiers du répértoires
    fs.readdir(pathString, async function (err, items) {
      const datas: DataDsn[] = []
      //On analyse les fichiers DSN 
      for (let i = 0; i < items.length; i++) {
        let dsnParser = new DsnParser()
        try {
          await dsnParser.asyncInit(pathString + '/' + items[i], { controleDsnVersion: true, deleteFile: true })
          console.log(dsnParser.contribution)
          let data: DataDsn = {
            dsn: dsnParser.dsn,
            establishment: dsnParser.establishment,
            mutual: dsnParser.mutual,
            employeeMutual: dsnParser.employeeMutual,
            contributionFund: dsnParser.contributionFund,
            employee: dsnParser.employee,
            workContract: dsnParser.workContract,
            contribution: dsnParser.contribution,
            base: dsnParser.base,
            rateAt: dsnParser.rateAt,
            rateMobility: dsnParser.rateMobility,
            //workStopping: dsnParser.workStopping
          }
          datas.push(data)
        } catch (e) {
          throw (e)
        }

      }
      //Etape 5 on va créer un fichier Exel
      const excelFileName = 'dsn.xlsx'
      await createExcelFile(pathString, excelFileName, datas)
      //Etape 6 on va charger le fichier Excel dans la mémoire
      const excelBuffer = fs.readFileSync(`${patch}/${excelFileName}`);
      //Etape 7 on supprime le répértoire
      removeDir(date.toString())
      //Etape 8 on répond au client
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${excelFileName}"`);
      res.setHeader('Accept-Ranges', 'bytes');
      res.setHeader('X-Powered-By', 'NextJs')

      return res.send(excelBuffer)
    });


  } catch (e) {
    //Etape 7 on supprime le répértoire
    res.status(500).json({ error: e })
  }

}

const createExcelFile = async (patch: string, fileName: string, datas: DataDsn[]) => {
  //Attention on utilise la méthode sync qui bloque JS voir pour passer sur l'API Async
  //Création du fichier
  const workbook = new Excel.Workbook();
  workbook.addWorksheet('DSN', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Etablissement', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Organismes sociaux', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Individus', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Contrat travail', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Affiliations', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Base', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Base assujeti', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Cotisations', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Taux AT', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Taux versement transport', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Absences', { properties: { tabColor: { argb: 'FFC0000' } } });


  for (let data of datas) {
    //Gestion de la feuille DSN
    const dsnSheet = workbook.getWorksheet('DSN')
    dsnSheet.columns = [
      { header: 'Mois', key: 'month', width: 10, outlineLevel: 1 },
      { header: 'Nom du logiciel', key: 'softwareName', width: 25 },
      { header: 'Fournisseur', key: 'provider', width: 10 },
      { header: 'Version du logiciel', key: 'softwareVersion', width: 10, outlineLevel: 1 },
      { header: 'type', key: 'type', width: 10, outlineLevel: 1 },
    ];
    dsnSheet.addRow({
      month: data.dsn.month,
      softwareName: data.dsn.softwareName,
      provider: data.dsn.provider,
      softwareVersion: data.dsn.softwareName,
      type: data.dsn.type,
    })
    //Gestion des établissements
    const establishmentSheet = workbook.getWorksheet('Etablissement')
    establishmentSheet.columns = [
      { header: 'Mois', key: 'month', width: 25 },
      { header: 'NIC', key: 'nic', width: 25 },
      { header: 'Code APET', key: 'apet', width: 25 },
      { header: 'Adresse', key: 'adress1', width: 25 },
      { header: 'Complément adresse', key: 'adress2', width: 25 },
      { header: 'Code postal', key: 'codeZip', width: 25 },

    ]
    for (let establishment of data.establishment) {
      establishmentSheet.addRow({
        month: data.dsn.month,
        nic: establishment.nic,
        apet: establishment.apet,
        adress1: establishment.adress1,
        adress2: establishment.adress2,
        codeZip: establishment.codeZip,
      })
    }

    //Gestion des OPS

    const contributionFundSheet = workbook.getWorksheet('Organismes sociaux')
    contributionFundSheet.columns = [
      { header: 'Mois', key: 'month', width: 25 },
      { header: 'Code DSN', key: 'codeDsn', width: 25 },
      { header: 'Organisme', key: 'name', width: 25 },
      { header: 'Adresse', key: 'adress1', width: 25 },
      { header: 'Code postal', key: 'codeZip', width: 25 },
      { header: 'Ville', key: 'city', width: 25 },
      { header: 'siret', key: 'siret', width: 25 },
    ]

    for (let contributionFund of data.contributionFund) {
      contributionFundSheet.addRow({
        month: data.dsn.month,
        codeDsn: contributionFund.codeDsn,
        name: contributionFund.name,
        adress1: contributionFund.adress1,
        codeZip: contributionFund.codeZip,
        city: contributionFund.city,
        siret: contributionFund.siret,
      })
    }
    //Gestion des individus
    const employeeSheet = workbook.getWorksheet('Individus')
    employeeSheet.columns = [
      { header: 'Mois', key: 'month', width: 25 },
      { header: 'Matricule', key: 'employeeId', width: 25 },
      { header: 'Numéro de Sécurité Sociale', key: 'numSS', width: 25 },
      { header: 'Département de naissance', key: 'codeZipBith', width: 25 },
      { header: 'Pays de naissance', key: 'countryBirth', width: 25 },
      { header: 'Nom', key: 'lastname', width: 25 },
      { header: 'Nom de famille', key: 'surname', width: 25 },
      { header: 'Prénom', key: 'firstname', width: 25 },
      { header: 'Sexe', key: 'sex', width: 25 },
      { header: 'Date anniversaire', key: 'birthday', width: 25 },
      { header: 'Lieu de naissance', key: 'placeOfBith', width: 25 },
      { header: 'Adresse', key: 'address1', width: 25 },
      { header: 'Complément de la localisation de la construction', key: 'address2', width: 25 },
      { header: 'Service de distribution, complément de localisation de la voie', key: 'address3', width: 25 },
      { header: 'Code postal', key: 'codeZip', width: 25 },
      { header: 'Ville', key: 'city', width: 25 },
      { header: 'Email', key: 'email', width: 25 },
      { header: 'Niveau etude', key: 'graduate', width: 25 },
      { header: `Niveau de diplôme préparé par l'individu`, key: 'v', width: 25 },

    ]
    for (let employee of data.employee) {
      employeeSheet.addRow({
        month: data.dsn.month,
        employeeId: employee.employeeId,
        numSS: employee.numSS,
        codeZipBith: employee.codeZipBith,
        countryBirth: employee.countryBirth,
        lastname: employee.lastname,
        surname: employee.surname,
        firstname: employee.firstname,
        sex: employee.sex,
        birthday: employee.birthday,
        placeOfBith: employee.placeOfBith,
        address1: employee.address1,
        address2: employee?.address2,
        address3: employee?.address3,
        codeZip: employee.codeZip,
        city: employee.city,
        email: employee.email,
        graduate: employee?.graduate,
        studies: employee?.studies
      })
    }

    //Gestion des contrats de travail
    const workContractSheet = workbook.getWorksheet('Contrat travail')
    workContractSheet.columns = [
      { header: 'Mois', key: 'month', width: 25 },
      { header: 'Matricule', key: 'employeeId', width: 25 },
      { header: 'Date début de contrat', key: 'startDate', width: 25 },
      { header: 'Date de fin prévisionnelle du contrat', key: 'endDate', width: 25 },
      { header: 'Statut du salarié (conventionnel)', key: 'status', width: 25 },
      { header: 'Code statut catégoriel Retraite Complémentaire obligatoire', key: 'retirement', width: 25 },
      { header: 'Code profession et catégorie socioprofessionnelle (PCS-ESE)', key: 'pcs', width: 25 },
      { header: 'Code complément PCS-ESE (pour la fonction publique : référentiels NEH, NET et grade de la NNE)', key: 'pcsBis', width: 25 },
      { header: `Libellé de l'emploi`, key: 'employmentLabel', width: 25 },
      { header: 'Nature du contrat', key: 'contract', width: 25 },
      { header: 'Dispositif de politique publique et conventionnel', key: 'publicDispPolitic', width: 25 },
      { header: 'Numéro du contrat', key: 'contractId', width: 25 },
      { header: 'Unité de mesure de la quotité de travail', key: 'DNACodeUnitTime', width: 25 },
      { header: `Quotité de travail de référence de l'entreprise pour la catégorie de salarié`, key: 'DSNWorkQuotaEstablishment', width: 25 },
      { header: `Quotité de travail du contrat`, key: 'DSNWorkQuotaWorkContract', width: 25 },
      { header: `Modalité d'exercice du temps de travail`, key: 'workTime', width: 25 },
      { header: `Complément de base au régime obligatoire`, key: 'ss', width: 25 },
      { header: `Code convention collective applicable`, key: 'idcc', width: 25 },
      { header: `Code régime de base risque maladie`, key: 'mal', width: 25 },
      { header: `Identifiant du lieu de travail`, key: 'estabWorkPlace', width: 25 },
      { header: `Code régime de base risque vieillesse`, key: 'vieillesse', width: 25 },
      { header: `Motif de recours`, key: 'pattern', width: 25 },
      { header: `Code caisse professionnelle de congés payés`, key: 'vacation', width: 25 },
      { header: `Taux de déduction forfaitaire spécifique pour frais professionnels`, key: 'rateProfessionalFess', width: 25 },
      { header: `Travailleur à l'étranger au sens du code de la Sécurité Sociale`, key: 'foreigner', width: 25 },
      { header: `Motif d'exclusion DSN`, key: 'exclusionDsn', width: 25 },
      { header: `Statut d'emploi du salarié`, key: 'statusEmployment', width: 25 },
      { header: `Code affectation Assurance chômage`, key: 'unemployment', width: 25 },
      { header: `Numéro interne employeur public`, key: 'idPublicEmployer', width: 25 },
      { header: `Type de gestion de l’Assurance chômage`, key: 'methodUnemployment', width: 25 },
      { header: `Date d'adhésion`, key: 'joiningDate', width: 25 },
      { header: `Date de dénonciation`, key: 'denunciationDate', width: 25 },
      { header: `Date d’effet de la convention de gestion`, key: 'dateManagementAgreement', width: 25 },
      { header: `Numéro de convention de gestion`, key: 'idAgreement', width: 25 },
      { header: `Code délégataire du risque maladie`, key: 'healthRiskDelegate', width: 25 },
      { header: `Code emplois multiples`, key: 'multipleJobCode', width: 25 },
      { header: `Code employeurs multiples`, key: 'multipleEmployerCode', width: 25 },
      { header: `Code régime de base risque accident du travail`, key: 'workAccidentRisk', width: 25 },
      { header: `Code risque accident du travail`, key: 'idWorkAccidentRisk', width: 25 },
      { header: `Positionnement dans la convention collective`, key: 'positionCollectiveAgreement', width: 25 },
      { header: `Code statut catégoriel APECITA`, key: 'apecita', width: 25 },
      { header: `Taux de cotisation accident du travail`, key: 'rateAt', width: 25 },
      { header: `Salarié à temps partiel cotisant à temps plein`, key: 'contributingFullTime', width: 25 },
      { header: `Rémunération au pourboire`, key: 'tip', width: 25 },
      { header: `Identifiant de l’établissement utilisateur`, key: 'useEstablishmentId', width: 25 },
      { header: `Numéro de label « Prestataire de services du spectacle vivant`, key: 'livePerfomances', width: 25 },
      { header: `Numéro de licence entrepreneur spectacle`, key: 'licences', width: 25 },
      { header: `Numéro objet spectacle`, key: 'showId', width: 25 },
      { header: `Statut organisateur spectacle`, key: 'showrunner', width: 25 },
      { header: `[FP] Code complément PCS-ESE pour la fonction publique d'Etat(emploi de la NNE)`, key: 'fpPcs', width: 25 },
      { header: `Nature du poste`, key: 'typePosition', width: 25 },
      { header: `[FP] Quotité de travail de référence de l'entreprise pour la catégorie de salarié dans l’hypothèse d’un poste à temps complet`, key: 'fpQuotite', width: 25 },
      { header: `Taux de travail à temps partiel`, key: 'partTimeWork', width: 25 },
      { header: `Code catégorie de service`, key: 'serviceCode', width: 25 },
      { header: `[FP] Indice brut`, key: 'fpIndice', width: 25 },
      { header: `[FP] Indice majoré`, key: 'fpIndiceMaj', width: 25 },
      { header: `[FP] Nouvelle bonification indiciaire (NBI)`, key: 'NBI', width: 25 },
      { header: `[FP] Indice brut d'origine`, key: 'indiceOriginal', width: 25 },
      { header: `[FP] Indice brut de cotisation dans un emploi supérieur (article 15)`, key: 'article15', width: 25 },
      { header: `[FP] Ancien employeur public`, key: 'oldEstablishment', width: 25 },
      { header: `[FP] Indice brut d’origine ancien salarié employeur public`, key: 'oldIndice', width: 25 },
      { header: `[FP] Indice brut d’origine sapeur-pompier professionnel (SPP)`, key: 'SPP', width: 25 },
      { header: `[FP] Maintien du traitement d'origine d'un contractuel titulaire`, key: 'contractual', width: 25 },
      { header: `[FP] Type de détachement`, key: 'secondment', width: 25 },
      { header: `Genre de navigation`, key: 'browsing', width: 25 },
      { header: `Taux de service actif`, key: 'activityDutyRate', width: 25 },
      { header: `Niveau de rémunération`, key: 'payLevel', width: 25 },
      { header: `Echelon`, key: 'echelon', width: 25 },
      { header: `Coefficient`, key: 'coefficient', width: 25 },
      { header: `Statut BOETH`, key: 'boeth', width: 25 },
      { header: `Complément de dispositif de politique publique`, key: 'addPublicPolicy', width: 25 },
      { header: `Cas de mise à disposition externe d'un individu de l'établissement`, key: 'arrangement', width: 25 },
      { header: `Catégorie de classement finale`, key: 'finaly', width: 25 },
      { header: `Identifiant du contrat d'engagement maritime`, key: 'navy', width: 25 },
      { header: `Collège (CNIEG)`, key: 'cnieg', width: 25 },
      { header: `Forme d'aménagement du temps de travail dans le cadre de l'activité partielle`, key: 'activityRate', width: 25 },
      { header: `Grade`, key: 'grade', width: 25 },
      { header: `[FP] Indice complément de traitement indiciaire (CTI)`, key: 'cti', width: 25 },
      { header: `FINESS géographique`, key: 'finess', width: 25 },

    ]

    for (let workContract of data.workContract) {
      workContractSheet.addRow({
        month: data.dsn.month,
        employeeId: workContract.employeeId,
        startDate: workContract.startDate,
        endDate: workContract?.contractEndDate,
        status: workContract?.status,
        retirement: workContract.retirement,
        pcs: workContract.pcs,
        pcsBis: workContract.pcsBis,
        employmentLabel: workContract.employmentLabel,
        contract: workContract.contract,
        publicDispPolitic: workContract.publicDispPolitic,
        contractId: workContract.contract,
        DNACodeUnitTime: workContract.DNACodeUnitTime,
        DSNWorkQuotaEstablishment: workContract.DSNWorkQuotaEstablishment,
        DSNWorkQuotaWorkContract: workContract.DSNWorkQuotaWorkContract,
        workTime: workContract.workTime,
        ss: workContract.ss,
        idcc: workContract.idcc,
        mal: workContract.mal,
        estabWorkPlace: workContract.estabWorkPlace,
        vieillesse: workContract.vieillesse,
        pattern: workContract.pattern,
        vacation: workContract.vacation,
        rateProfessionalFess: workContract?.rateProfessionalFess,
        foreigner: workContract?.foreigner,
        exclusionDsn: workContract?.exclusionDsn,
        statusEmployment: workContract.statusEmployment,
        unemployment: workContract.unemployment,
        idPublicEmployer: workContract.idPublicEmployer,
        methodUnemployment: workContract.methodUnemployment,
        joiningDate: workContract.joiningDate,
        denunciationDate: workContract.denunciationDate,
        dateManagementAgreement: workContract.dateManagementAgreement,
        idAgreement: workContract.idAgreement,
        healthRiskDelegate: workContract.healthRiskDelegate,
        multipleJobCode: workContract.multipleJobCode,
        multipleEmployerCode: workContract.multipleEmployerCode,
        workAccidentRisk: workContract.workAccidentRisk,
        idWorkAccidentRisk: workContract.idWorkAccidentRisk,
        positionCollectiveAgreement: workContract.positionCollectiveAgreement,
        apecita: workContract.apecita,
        rateAt: workContract.rateAt,
        contributingFullTime: workContract.contributingFullTime,
        tip: workContract.tip,
        useEstablishmentId: workContract.useEstablishmentId,
        livePerfomances: workContract?.livePerfomances,
        licences: workContract?.licences,
        showId: workContract?.showId,
        showrunner: workContract?.showrunner,
        fpPcs: workContract?.fpPcs,
        typePosition: workContract?.typePosition,
        fpQuotite: workContract?.fpQuotite,
        partTimeWork: workContract?.partTimeWork,
        serviceCode: workContract?.serviceCode,
        fpIndice: workContract?.fpIndice,
        fpIndiceMaj: workContract?.fpIndiceMaj,
        NBI: workContract?.NBI,
        indiceOriginal: workContract?.indiceOriginal,
        article15: workContract?.article15,
        oldEstablishment: workContract?.oldEstablishment,
        oldIndice: workContract?.oldIndice,
        SPP: workContract?.SPP,
        contractual: workContract?.contractual,
        secondment: workContract?.secondment,
        browsing: workContract?.browsing,
        activityDutyRate: workContract?.activityDutyRate,
        payLevel: workContract?.payLevel,
        echelon: workContract?.echelon,
        coefficient: workContract?.coefficient,
        boeth: workContract?.boeth,
        addPublicPolicy: workContract?.addPublicPolicy,
        arrangement: workContract?.arrangement,
        finaly: workContract?.finaly,
        navy: workContract?.navy,
        cnieg: workContract?.cnieg,
        activityRate: workContract?.activityRate,
        grade: workContract?.grade,
        cti: workContract?.cti,
        finess: workContract?.finess
      })
    }
    //Gestion des bases 
    const baseSheet = workbook.getWorksheet('Base')
    baseSheet.columns = [
      { header: 'Mois', key: 'month', width: 25 },
      { header: 'Matricule', key: 'employeeId', width: 25 },
      { header: 'Code de base assujettie', key: 'idBase', width: 25 },
      { header: 'Date de début de période de rattachement', key: 'startDate', width: 25 },
      { header: 'Date de fin de période de rattachement', key: 'endDate', width: 25 },
      { header: 'Montant', key: 'amount', width: 25 },
      { header: 'Identifiant technique Affiliation', key: 'idTechAff', width: 25 },
      { header: 'Numéro du contrat', key: 'idContract', width: 25 },
      { header: 'CRM', key: 'crm', width: 25 },

    ]
    for (let base of data.base) {
      baseSheet.addRow({
        month: base.date,
        employeeId: base.employeeId,
        idBase: base.idBase,
        startDate: base.startDate,
        endDate: base.endDate,
        amount: base.amount,
        idTechAff: base?.idTechAff,
        idContract: base?.idContract,
        crm: base?.crm

      })
    }

    //Gestion des cotisations
    const contributionSheet = workbook.getWorksheet('Cotisations')
    contributionSheet.columns = [
      { header: 'Mois', key: 'month', width: 25 },
      { header: 'Matricule', key: 'employeeId', width: 25 },
      { header: 'Code de cotisation', key: 'idContribution', width: 25 },
      { header: 'Identifiant Organisme de Protection Sociale', key: 'ops', width: 25 },
      { header: `Montant d assiette`, key: 'baseContribution', width: 25 },
      { header: `Montant de cotisation`, key: 'amountContribution', width: 25 },
      { header: `Code INSEE commune`, key: 'idInsee', width: 25 },
      { header: `Identifiant du CRM à l origine de la régularisation`, key: 'crmContribution', width: 25 },
      { header: `Taux de cotisation`, key: 'rateContribution', width: 25 },

    ]
    for (let contribution of data.contribution) {
      if (contribution.amountContribution) {
        contributionSheet.addRow({
          month: contribution.date,
          employeeId: contribution.employeeId,
          idContribution: contribution.idContribution,
          ops: contribution?.ops,
          baseContribution: contribution?.baseContribution,
          amountContribution: contribution.amountContribution,
          idInsee: contribution?.idInsee,
          crmContribution: contribution?.crmContribution,
          rateContribution: contribution?.rateContribution
        })
      }

    }

    //Gestion des taux AT

    const rateAtSheet = workbook.getWorksheet('Taux AT')
    rateAtSheet.columns = [
      { header: 'Mois', key: 'month', width: 25 },
      { header: 'Siret', key: 'siret', width: 25 },
      { header: 'Code risque', key: 'code', width: 25 },
      { header: 'Taux', key: 'rate', width: 25 },

    ]
    for (let rateAT of data.rateAt) {
      rateAtSheet.addRow({
        month: rateAT?.date,
        siret: rateAT.siret,
        code: rateAT.code,
        rate: rateAT.rate
      })
    }

    /** 
    //Gestion des taux versement transport

    const rateMobilitySheet = workbook.getWorksheet('Taux versement transport')
    rateMobilitySheet.columns = [
      { header: 'Mois', key: 'month', width: 25 },
      { header: 'Siret', key: 'siret', width: 25 },
      { header: 'Code insee', key: 'codeInsee', width: 25 },
      { header: 'Taux', key: 'rate', width: 25 },
    ]

    for (let rateMobility of data.rateMobility) {
      rateMobilitySheet.addRow = ({
        codeInsee: rateMobility.insee,
        rate: rateMobility.rate
      })
    }
*/
    /** 
    //Gestion des absences 

    const workStoppingSheet = workbook.getWorksheet('Absences')
    workStoppingSheet.columns = [
      { header: 'Mois', key: 'month', width: 25 },
      { header: 'Siret', key: 'siret', width: 25 },
      { header: `Matricule`, key: 'employeeId', width: 25 },
      { header: `Motif de l'arrêt`, key: 'reasonStop', width: 25 },
      { header: 'Date du dernier jour travaillé', key: 'lastDayWorked', width: 25 },
      { header: 'Date de fin prévisionnelle', key: 'estimatedEndDate', width: 25 },
      { header: 'Subrogation', key: 'subrogation', width: 25 },
      { header: 'Date de début de subrogation', key: 'subrogationStartDate', width: 25 },
      { header: 'Date de début de subrogation', key: 'subrogationEndDate', width: 25 },
      { header: 'IBAN', key: 'iban', width: 25 },
      { header: 'BIC', key: 'bic', width: 25 },
      { header: 'Date de la reprise', key: 'recoveryDate', width: 25 },
      { header: 'Motif de la reprise', key: 'reasonRecovery', width: 25 },
      { header: `Date de l'accident ou de la première constatation`, key: 'dateWorkAccident', width: 25 },
      { header: `SIRET Centralisateur`, key: 'SIRETCentralizer', width: 25 },

    ]

    for (let workStopping of data.workStopping) {
      workStoppingSheet.addRow({
        month: workStopping.date,
        siret: workStopping.siret,
        employeeId: workStopping.employeeId,
        reasonStop: workStopping.reasonStop,
        lastDayWorked: workStopping.lastDayWorked,
        estimatedEndDate: workStopping?.estimatedEndDate,
        subrogation: workStopping?.subrogation,
        subrogationStartDate: workStopping?.subrogationStartDate,
        subrogationEndDate: workStopping?.subrogationEndDate,
        iban: workStopping?.iban,
        bic: workStopping?.bic,
        recoveryDate: workStopping?.recoveryDate,
        reasonRecovery: workStopping?.reasonRecovery,
        dateWorkAccident: workStopping?.dateWorkAccident,
        SIRETCentralizer: workStopping?.SIRETCentralizer
      })
    }
    */
    //Ecriture du fichier
    await workbook.xlsx.writeFile(`${patch}/${fileName}`);
  }

}
