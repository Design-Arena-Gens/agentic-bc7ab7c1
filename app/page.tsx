"use client";

import { useMemo, useState } from "react";

type OutlookPathKey = "m365" | "office2019" | "office2016" | "custom";

const OUTLOOK_PRESETS: Record<
  Exclude<OutlookPathKey, "custom">,
  { label: string; path: string }
> = {
  m365: {
    label: "Microsoft 365 (Office Click-to-Run)",
    path: "C:\\\\Program Files\\\\Microsoft Office\\\\root\\\\Office16\\\\OUTLOOK.EXE"
  },
  office2019: {
    label: "Office 2019 (64 bits)",
    path: "C:\\\\Program Files\\\\Microsoft Office\\\\Office16\\\\OUTLOOK.EXE"
  },
  office2016: {
    label: "Office 2016/2013 (32 bits)",
    path: "C:\\\\Program Files (x86)\\\\Microsoft Office\\\\Office16\\\\OUTLOOK.EXE"
  }
};

const DEFAULT_MESSAGE = [
  "Bonjour,",
  "",
  "Ceci est un rappel automatique envoyé à chaque démarrage de l'ordinateur.",
  "",
  "Bonne journée !"
].join("\n");

const sanitizeForPowerShell = (value: string) =>
  value.replace(/`/g, "``").replace(/\$/g, "`$").replace(/"/g, '`"');

export default function Home() {
  const [recipient, setRecipient] = useState("votre.adresse@email.com");
  const [subject, setSubject] = useState("Rappel automatique");
  const [message, setMessage] = useState(DEFAULT_MESSAGE);
  const [startupDelay, setStartupDelay] = useState(12);
  const [keepOpen, setKeepOpen] = useState(true);
  const [outlookPathKey, setOutlookPathKey] = useState<OutlookPathKey>("m365");
  const [customPath, setCustomPath] = useState("C:\\\\Program Files\\\\Microsoft Office\\\\root\\\\Office16\\\\OUTLOOK.EXE");
  const [copied, setCopied] = useState(false);

  const script = useMemo(() => {
    const selectedPath =
      outlookPathKey === "custom" ? customPath : OUTLOOK_PRESETS[outlookPathKey].path;
    const sanitizedSubject = sanitizeForPowerShell(subject);
    const sanitizedRecipient = sanitizeForPowerShell(recipient);
    const hereString = message.endsWith("\n") ? `${message}` : `${message}\n`;
    const bodySection = `@"
${hereString}"@`;

    const closeSnippet = keepOpen
      ? "# Outlook reste ouvert pour que vous puissiez continuer à l'utiliser."
      : [
          'Write-Verbose "Fermeture d\'Outlook..."',
          "$outlookProcess = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue",
          "if ($outlookProcess) {",
          "    $outlookProcess.CloseMainWindow() | Out-Null",
          "}"
        ].join("\n");

    return [
      "# Automatise l'ouverture d'Outlook et l'envoi d'un email au démarrage",
      "param(",
      '    [string]$OutlookPath = "' + selectedPath + '"',
      ")",
      "",
      "$ErrorActionPreference = \"Stop\"",
      "",
      "Write-Verbose \"Vérification de la présence d'Outlook...\"",
      "if (-not (Test-Path $OutlookPath)) {",
      "    throw \"Outlook n'a pas été trouvé à l'emplacement spécifié : $OutlookPath\"",
      "}",
      "",
      "if (-not (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)) {",
      "    Write-Verbose \"Lancement d'Outlook...\"",
      "    Start-Process -FilePath $OutlookPath",
      `    Start-Sleep -Seconds ${Math.max(0, startupDelay)}`,
      "} else {",
      "    Write-Verbose \"Outlook est déjà en cours d'exécution.\"",
      "}",
      "",
      "Write-Verbose \"Connexion à Outlook...\"",
      "$outlook = New-Object -ComObject Outlook.Application",
      "$namespace = $outlook.GetNamespace(\"MAPI\")",
      "$namespace.Logon()",
      "",
      "Write-Verbose \"Création du message...\"",
      "$mail = $outlook.CreateItem(0)",
      `$mail.To = "${sanitizedRecipient}"`,
      `$mail.Subject = "${sanitizedSubject}"`,
      "$mail.Body = " + bodySection,
      "",
      "Write-Verbose \"Envoi du message...\"",
      "$mail.Send()",
      "",
      closeSnippet,
      "",
      "Write-Verbose \"Terminé.\""
    ].join("\n");
  }, [customPath, keepOpen, message, outlookPathKey, recipient, startupDelay, subject]);

  const handleCopy = async () => {
    await navigator.clipboard.writeText(script);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <main className="container">
      <h1>Automatiser Outlook au démarrage</h1>
      <p className="tip">
        Configurez un script PowerShell qui ouvre Outlook à chaque démarrage et vous envoie un email
        personnalisé.
      </p>
      <div className="grid">
        <section className="panel">
          <h2>Paramètres du message</h2>
          <div className="inline-field">
            <label htmlFor="recipient">Destinataire</label>
            <input
              id="recipient"
              value={recipient}
              onChange={(event) => setRecipient(event.target.value)}
              placeholder="prenom.nom@domaine.com"
              spellCheck={false}
            />
          </div>
          <div className="inline-field">
            <label htmlFor="subject">Objet</label>
            <input
              id="subject"
              value={subject}
              onChange={(event) => setSubject(event.target.value)}
            />
          </div>
          <div className="inline-field">
            <label htmlFor="message">Message</label>
            <textarea
              id="message"
              value={message}
              onChange={(event) => setMessage(event.target.value)}
              rows={6}
            />
          </div>
          <div className="inline-field">
            <label htmlFor="delay">Attente après le lancement (secondes)</label>
            <input
              id="delay"
              type="number"
              min={0}
              max={120}
              value={startupDelay}
              onChange={(event) => setStartupDelay(Number(event.target.value))}
            />
          </div>
          <div className="inline-field">
            <label>Gestion d&apos;Outlook</label>
            <div className="radio-group">
              <button
                type="button"
                className="radio-pill"
                data-active={keepOpen}
                onClick={() => setKeepOpen(true)}
              >
                Laisser Outlook ouvert
              </button>
              <button
                type="button"
                className="radio-pill"
                data-active={!keepOpen}
                onClick={() => setKeepOpen(false)}
              >
                Fermer après envoi
              </button>
            </div>
          </div>
        </section>

        <section className="panel">
          <h2>Chemin d&apos;Outlook</h2>
          <div className="grid">
            {(Object.keys(OUTLOOK_PRESETS) as Array<Exclude<OutlookPathKey, "custom">>).map(
              (key) => (
                <button
                  key={key}
                  type="button"
                  className="radio-pill"
                  data-active={outlookPathKey === key}
                  onClick={() => setOutlookPathKey(key)}
                >
                  {OUTLOOK_PRESETS[key].label}
                </button>
              )
            )}
            <button
              type="button"
              className="radio-pill"
              data-active={outlookPathKey === "custom"}
              onClick={() => setOutlookPathKey("custom")}
            >
              Chemin personnalisé
            </button>
          </div>
          {outlookPathKey === "custom" ? (
            <div className="inline-field">
              <label htmlFor="customPath">Chemin complet</label>
              <input
                id="customPath"
                value={customPath}
                onChange={(event) => setCustomPath(event.target.value)}
              />
            </div>
          ) : null}
          <p className="tip">
            Vérifiez le chemin via Outlook → Fichier → Options → Compléments → bouton{" "}
            <code>Aller...</code>, ou en faisant clic droit sur le raccourci Outlook puis « Ouvrir
            l&apos;emplacement du fichier ».
          </p>
        </section>

        <section className="panel">
          <div className="code-toolbar">
            <span className="pill">Script PowerShell</span>
            <button type="button" onClick={handleCopy}>
              {copied ? "Copié !" : "Copier"}
            </button>
          </div>
          <pre className="code-block">{script}</pre>
        </section>

        <section className="panel">
          <h2>Étapes d&apos;installation au démarrage</h2>
          <ol>
            <li>
              Ouvrez le Bloc-notes, collez le script ci-dessus puis enregistrez-le sous le nom{" "}
              <code>Outlook-AutoStart.ps1</code> dans votre dossier Documents.
            </li>
            <li>
              Cliquez sur Démarrer → tapez <strong>Planificateur de tâches</strong> → créez une
              nouvelle tâche basique.
            </li>
            <li>
              Choisissez le déclencheur <em>À l&apos;ouverture de session</em> (ou <em>Au démarrage</em>{" "}
              si vous préférez).
            </li>
            <li>
              Dans l&apos;action, sélectionnez <em>Démarrer un programme</em> puis saisissez{" "}
              <code>powershell.exe</code> comme programme. Dans « Ajouter des arguments », indiquez{" "}
              <code>-ExecutionPolicy Bypass -File &quot;C:\Users\&lt;vous&gt;\Documents\Outlook-AutoStart.ps1&quot;</code>.
            </li>
            <li>
              Validez, puis redémarrez l&apos;ordinateur pour tester. Vous recevrez l&apos;email et
              Outlook sera ouvert automatiquement.
            </li>
          </ol>
        </section>
      </div>
      <p className="footer">
        Conseil : gardez le script dans un emplacement accessible et ajustez-le si votre adresse,
        l&apos;objet ou le chemin d&apos;Outlook change.
      </p>
    </main>
  );
}
