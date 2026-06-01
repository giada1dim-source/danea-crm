# Danea CRM Agenti - versione Supabase

## Avvio locale

```cmd
npm install
npm run dev
```

Apri http://localhost:3000

## Collegamento Supabase

1. In Supabase apri **SQL Editor**.
2. Apri il file `supabase/schema.sql`, copia tutto il contenuto e premi **Run** in Supabase.
3. In Supabase vai su **Project Settings > API**.
4. Copia:
   - Project URL
   - anon public key
5. Nella cartella del progetto copia `.env.local.example` e rinominalo in `.env.local`.
6. Incolla i valori:

```env
NEXT_PUBLIC_SUPABASE_URL=...
NEXT_PUBLIC_SUPABASE_ANON_KEY=...
```

7. Ferma il server con CTRL+C e riavvia:

```cmd
npm run dev
```

## Uso

1. Carica `clienti danea.xlsx`.
2. Carica `fatture.DefXml`.
3. Premi **Salva DB**.
4. Da un altro PC/browser premi **Carica DB** per recuperare i dati salvati.

## Sicurezza

Questa è una versione MVP. Le policy Supabase sono aperte per testare velocemente. Prima di darla agli agenti, bisogna aggiungere login e permessi admin/agente.
