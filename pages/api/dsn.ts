import formidable from "formidable";
import type { NextApiRequest, NextApiResponse } from 'next'
import path from "path";
//import fs from "fs/promises";
import fs from 'fs';
import { DsnParser } from "@fibre44/dsn-parser";
import Excel from 'exceljs'
type Data = {
  name?: string,
  error?: string | unknown,
  succes?: string
} | Buffer
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
    //Etape 4 on va créer un fichier Exel
    const excelFileName = 'dsn.xlsx'
    await createExcelFile(pathString, excelFileName)

    //Etape 5 on liste les fichiers du répértoires
    fs.readdir(pathString, async function (err, items) {
      for (let i = 0; i < items.length; i++) {
        let dsnParser = new DsnParser()
        try {
          await dsnParser.init(pathString + '/' + items[i], { controleDsnVersion: true, deleteFile: true })

        } catch (e) {
          throw (e)
        }

      }
    });
    //Etape 6 on va charger le fichier Excel dans la mémoire
    const excelBuffer = fs.readFileSync(`${patch}/${excelFileName}`);
    //Etape 7 on supprime le répértoire
    removeDir(date.toString())
    //Etape 8 on répond au client
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="test.xlsx"');
    res.setHeader('Accept-Ranges', 'bytes');
    res.setHeader('X-Powered-By', 'NextJs')

    return res.send(excelBuffer)
  } catch (e) {
    res.status(500).json({ error: e })
  }

}

const createExcelFile = async (patch: string, fileName: string,) => {

  const workbook = new Excel.Workbook();

  workbook.addWorksheet('DSN', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Etablissement', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Organismes_sociaux', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Individus', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Contrat_travail', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Affiliations', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Cotisations', { properties: { tabColor: { argb: 'FFC0000' } } });
  await workbook.xlsx.writeFile(`${patch}/${fileName}`);

}
