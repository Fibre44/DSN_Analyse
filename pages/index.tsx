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
      <main className="container">
        <form ref={form} onSubmit={handleSubmit} encType='multipart/form-data'>
          <div className="form-group">
            <label htmlFor="dsn">Selectionner vos fichiers DSN</label>
            <input type="file" name='dsn' className="form-control" id="dsn" accept=".dsn,.txt" multiple required />
          </div>

          <button type='submit' className="btn btn-primary" >Envoyer les données</button>
        </form>

      </main>
    </>
  )
}
