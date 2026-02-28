# Handleiding voor Publicatie (Hosting)

Om deze applicatie permanent beschikbaar te maken op het internet (zodat je hem op je telefoon of tablet kunt gebruiken), kun je gebruik maken van gratis hostingdiensten zoals **Vercel** of **Netlify**.

Hieronder staan de stappen om dit te doen.

## Optie 1: Via Vercel (Aanbevolen - Eenvoudigst)

Vercel is geoptimaliseerd voor dit soort applicaties en is gratis voor persoonlijk gebruik.

1.  **Download de code**:
    *   Klik in de editor op de knop om de bestanden te downloaden (indien beschikbaar) of kopieer de bestanden naar een map op je computer.
    *   Zorg dat je alle bestanden hebt, inclusief `package.json`, `vite.config.ts`, `index.html` en de `src` map.

2.  **Maak een GitHub account aan** (als je die nog niet hebt) op [github.com](https://github.com).
3.  **Upload de code naar GitHub**:
    *   Maak een nieuwe 'Repository' aan.
    *   Upload de bestanden hiernaartoe.

4.  **Ga naar Vercel**:
    *   Ga naar [vercel.com](https://vercel.com) en log in met je GitHub account.
    *   Klik op "Add New..." -> "Project".
    *   Selecteer de Repository die je net hebt aangemaakt.
    *   Klik op "Deploy".

Vercel herkent automatisch dat het een Vite/React applicatie is en zet deze online. Je krijgt binnen enkele minuten een link (bijv. `medset-loaner.vercel.app`).

## Optie 2: Handmatig Bouwen (Voor eigen server)

Als je een eigen webserver hebt of de bestanden ergens anders wilt neerzetten:

1.  Open de terminal in de projectmap.
2.  Installeer de benodigdheden (als je Node.js hebt geïnstalleerd):
    ```bash
    npm install
    ```
3.  Bouw de applicatie:
    ```bash
    npm run build
    ```
4.  Er verschijnt nu een map `dist` in je project.
5.  De inhoud van deze `dist` map zijn de kant-en-klare bestanden (HTML, CSS, JS). Deze kun je uploaden naar elke willekeurige webhosting via FTP of een controlepaneel.

## Gebruik op Mobiel (App-ervaring)

Zodra je applicatie online staat via een van de bovenstaande methoden:

1.  Open de link op je iPhone of Android toestel in de browser (Safari of Chrome).
2.  Tik op de "Delen" knop (iPhone) of het menu met drie puntjes (Android).
3.  Kies **"Zet op beginscherm"** (Add to Home Screen).
4.  De applicatie verschijnt nu als een icoontje op je telefoon en opent zonder browserbalken, net als een echte app.
