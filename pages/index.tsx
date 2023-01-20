import Head from 'next/head'
import { useRef } from 'react';


export default function Home() {
  const form = useRef(null)
  const handleSubmit = async (e: React.ChangeEvent<HTMLFormElement>) => {
    e.preventDefault()
    if (form.current) {
      const formData = new FormData(form.current)
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
            <label htmlFor="">Selectionner vos fichiers DSN</label>
            <input type="file" className="form-control" id="dsn" />
          </div>

          <button type='submit'>Envoyer les fichiers</button>
        </form>

      </main>
    </>
  )
}
