import './globals.css';

export const metadata = {
  title: 'Danea CRM Agenti',
  description: 'Dashboard clienti, fatture, agenti e visite collegata a export Danea Easyfatt.'
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return <html lang="it"><body>{children}</body></html>;
}
