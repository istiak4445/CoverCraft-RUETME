import type {Metadata} from 'next';
import './globals.css';
export const metadata:Metadata={title:'CoverCraft — RUET Mechanical',description:'Editable A4 assignment and lab report cover page generator for RUET Mechanical Engineering.',openGraph:{title:'CoverCraft — RUET Mechanical',description:'Editable A4 assignment & lab report covers',images:['/og.png']},twitter:{card:'summary_large_image',title:'CoverCraft — RUET Mechanical',description:'Editable A4 assignment & lab report covers',images:['/og.png']}};
export default function RootLayout({children}:{children:React.ReactNode}){return <html lang="en"><body>{children}</body></html>}
