import Head from 'next/head'
import { useRef } from 'react';


export default function Home() {
  const form = useRef(null)
  const handleSubmit = async (e: React.ChangeEvent<HTMLFormElement>) => {
    e.preventDefault()
    console.log(e)
    if (form.current) {
      const formData = new FormData(form.current)
      console.log(formData)
      const response = await fetch('/api/dsn', {
        method: 'POST',
        body: formData,
      });
      if (response.ok) {
        const file = window.URL.createObjectURL(await response.blob());
        window.location.assign(file);

      } else {
        console.error('ko')
      }
    }
  }
  return (
    <>
      <Head>
        <title>Extraction DSN vers Excel</title>
        <meta name="description" content="Extraction DSN vers Excel" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <link rel="icon" href="/favicon.ico" />
      </Head>
      <main className="container w-50 p-3">
        <form ref={form} onSubmit={handleSubmit} encType='multipart/form-data' className='form-signin'>
          <div className='mb-4'>
            <h1 className='h3 mb-3 font-weight-normal'>Outils d &apos; export des données DSN vers Excel</h1>
            <p className='text-justify'>Ce service permet d &apos;exporter les données de vos DSN vers un fichier Excel. L&apos;outil peut prendre en charge X fichiers DSN sur X périodes.</p>
            <p>Liste des informations exportables : </p>
            <ul>
              <li>DSN information</li>
              <li>Etablissements</li>
              <li>Liste des organismes de protection sociale</li>
              <li>Liste des libellés d&apos;emploi</li>
              <li>Salariés</li>
              <li>Contrat de travail</li>
              <li>Contrat complémentaires</li>
              <li>Liste des affiliations des salariés</li>
              <li>Bases des cotisations des salariés</li>
              <li>Cotisations des salariés</li>
            </ul>
          </div>
          <div className="form-group mb-4">
            <label htmlFor="dsn">Selectionner vos fichiers DSN</label>
            <input type="file" name='dsn' className="form-control" id="dsn" accept=".dsn,.txt" multiple required />
          </div>

          <button type='submit' className="btn btn-lg btn-primary btn-block" >Envoyer les données</button>
        </form>
      </main>
    </>
  )
}
