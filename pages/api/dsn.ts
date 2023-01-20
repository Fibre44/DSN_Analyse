import formidable from "formidable";
import type { NextApiRequest, NextApiResponse } from 'next'
import path from "path";
//import fs from "fs/promises";
import fs from 'fs';
import { DsnParser } from "@fibre44/dsn-parser";
import Excel from 'exceljs'
import { contributionFundObject, dsnObject, EmployeeObject, establishmentObject, mutualEmployeeObject, mutualObject, WorkContractObject } from "@fibre44/dsn-parser/lib/dsn";
type Data = {
  name?: string,
  error?: string | unknown,
  succes?: string
} | Buffer

type DataDsn = {
  dsn: dsnObject,
  establishment: establishmentObject[],
  mutualEmployee: mutualEmployeeObject[]
  mutual: mutualObject[],
  contributionFund: contributionFundObject[],
  employee: EmployeeObject[],
  workContract: WorkContractObject[]
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
          await dsnParser.init(pathString + '/' + items[i], { controleDsnVersion: true, deleteFile: true })

          let data: DataDsn = {
            dsn: dsnParser.dsn,
            establishment: dsnParser.establishment,
            mutualEmployee: dsnParser.employeeMutual,
            mutual: dsnParser.mutual,
            contributionFund: dsnParser.contributionFund,
            employee: dsnParser.employee,
            workContract: dsnParser.workContract
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
    res.status(500).json({ error: e })
  }

}

const createExcelFile = async (patch: string, fileName: string, datas: DataDsn[]) => {
  //Attention on utilise la méthode sync qui bloque JS voir pour passer sur l'API Async
  //Création du fichier
  const workbook = new Excel.Workbook();
  workbook.addWorksheet('DSN', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Etablissement', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Organismes_sociaux', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Individus', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Contrat_travail', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Affiliations', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Cotisations', { properties: { tabColor: { argb: 'FFC0000' } } });
  for (let data of datas) {
    //Gestion de la feuille DSN
    const dsnWorkboox = workbook.getWorksheet('DSN')
    dsnWorkboox.columns = [
      { header: 'Nom du logiciel', key: 'softwareName', width: 25 },
      { header: 'Fournisseur', key: 'provider', width: 10 },
      { header: 'Version du logiciel', key: 'softwareVersion', width: 10, outlineLevel: 1 },
      { header: 'type', key: 'type', width: 10, outlineLevel: 1 },
      { header: 'Mois', key: 'month', width: 10, outlineLevel: 1 },
    ];
    dsnWorkboox.addRow({
      softwareName: data.dsn.softwareName,
      provider: data.dsn.provider,
      softwareVersion: data.dsn.softwareName,
      type: data.dsn.type,
      month: data.dsn.month
    })
    //Gestion des établissements
    const establishmentWorkboox = workbook.getWorksheet('Etablissement')
    establishmentWorkboox.columns = [
      { header: 'NIC', key: 'nic', width: 25 },
      { header: 'Code APET', key: 'apet', width: 25 },
      { header: 'Adresse', key: 'adress1', width: 25 },
      { header: 'Complément adresse', key: 'adress2', width: 25 },
      { header: 'Code postal', key: 'codeZip', width: 25 },
      { header: 'Mois', key: 'month', width: 25 },

    ]
    for (let establishment of data.establishment) {
      establishmentWorkboox.addRow({
        nic: establishment.nic,
        apet: establishment.apet,
        adress1: establishment.adress1,
        adress2: establishment.adress2,
        codeZip: establishment.codeZip,
        month: data.dsn.month
      })
    }

    //Gestion des individus
    const employeeWorkboox = workbook.getWorksheet('Individus')
    employeeWorkboox.columns = [
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
      { header: 'Code postal', key: 'codeZip', width: 25 },
      { header: 'Ville', key: 'city', width: 25 },
      { header: 'Email', key: 'email', width: 25 },
    ]
    for (let employee of data.employee) {
      employeeWorkboox.addRow({
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
        codeZip: employee.codeZip,
        city: employee.city,
        email: employee.email
      })
    }

  }


  //Ecriture du fichier
  await workbook.xlsx.writeFile(`${patch}/${fileName}`);

}
