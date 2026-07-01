import { mkdir, unlink, writeFile } from 'node:fs/promises'
import path from 'node:path'
import type { marpCli as MarpCliFn } from '@marp-team/marp-cli'
import { nanoid } from 'nanoid'
import {
  commands,
  ExtensionContext,
  ProgressLocation,
  Uri,
  window,
  workspace,
} from 'vscode'
import {
  detectBrowserPath,
  generateNativePptx,
} from '@monsieurpablo/marp-native-pptx'

export function activate(context: ExtensionContext) {
  context.subscriptions.push(
    commands.registerCommand('marpToEditablePptx.export', exportCommand),
  )
}

export function deactivate() {
  // no-op
}

async function exportCommand(): Promise<void> {
  const editor = window.activeTextEditor
  if (!editor) {
    window.showErrorMessage('No active Markdown file.')
    return
  }

  const doc = editor.document
  if (doc.languageId !== 'markdown') {
    window.showErrorMessage('The active file is not a Markdown file.')
    return
  }

  if (doc.uri.scheme !== 'file') {
    window.showErrorMessage(
      'Please save the file to a local folder before exporting.',
    )
    return
  }

  if (doc.isDirty) {
    const answer = await window.showWarningMessage(
      'The file has unsaved changes. Save before exporting?',
      { modal: true },
      'Save and Export',
    )
    if (answer !== 'Save and Export') return
    await doc.save()
  }

  const defaultUri = Uri.file(doc.uri.fsPath.replace(/\.md$/i, '.pptx'))
  const saveUri = await window.showSaveDialog({
    defaultUri,
    filters: { PowerPoint: ['pptx'] },
    title: 'Export to Editable PPTX',
  })
  if (!saveUri) return

  await window.withProgress(
    {
      location: ProgressLocation.Notification,
      title: 'Exporting editable PPTX…',
      cancellable: false,
    },
    async () => {
      const tmpId = nanoid()
      // Place the temporary HTML next to the source Markdown so that
      // marp-cli resolves relative image paths (e.g. .attachments/) from
      // the correct directory. Using os.tmpdir() breaks relative paths
      // because marp-cli resolves media relative to the HTML output location.
      const htmlTmpPath = path.join(
        path.dirname(doc.uri.fsPath),
        `.marp-editable-pptx-${tmpId}.html`,
      )

      try {
        // Step 1: Convert Markdown → HTML via @marp-team/marp-cli
        const { marpCli } = (await import('@marp-team/marp-cli')) as {
          marpCli: typeof MarpCliFn
        }

        // Forward --html when the user has set markdown.marp.html to 'all'.
        // This preserves <script> tags in the marp-cli HTML output, which is
        // required for runtime rendering (e.g. mermaid.js via div.mermaid).
        // Matches the same logic used by marp-vscode's marpCoreOptionForCLI.
        const htmlSetting = workspace
          .getConfiguration('markdown.marp')
          .get<string>('html')
        const marpCliArgs = [
          doc.uri.fsPath,
          '-o',
          htmlTmpPath,
          '--allow-local-files',
          ...(htmlSetting === 'all' ? ['--html'] : []),
        ]

        // Marp CLI discovers .marprc.yml / themeSet via cosmiconfig starting
        // from process.cwd(). In the extension host cwd is not the workspace,
        // so custom themes and local config are silently ignored (issue #19).
        // Run the conversion from the Markdown file's directory so config
        // resolves exactly as `marp` on the CLI would (cosmiconfig also walks
        // up to the workspace root from there).
        // ponytail: process-global chdir; safe because the export is awaited
        // sequentially. Replace with a cwd option if marp-cli ever exposes one
        // (its programmatic API currently does not).
        const prevCwd = process.cwd()
        process.chdir(path.dirname(doc.uri.fsPath))
        let exitCode: number
        try {
          exitCode = await marpCli(marpCliArgs, {})
        } finally {
          process.chdir(prevCwd)
        }

        if (exitCode !== 0) {
          throw new Error(`Marp CLI exited with code ${exitCode}`)
        }

        // Step 2: Detect Chromium browser
        const browserPath = detectBrowserPath('auto', undefined)
        if (!browserPath) {
          throw new Error(
            'Could not find a Chromium-based browser required for PPTX export. ' +
              'Please install Google Chrome or Microsoft Edge.',
          )
        }

        // Step 3: Generate editable PPTX from HTML
        const pptxBuffer = await generateNativePptx({
          htmlPath: htmlTmpPath,
          browserPath,
        })

        // Step 4: Write output
        await mkdir(path.dirname(saveUri.fsPath), { recursive: true })
        await writeFile(saveUri.fsPath, pptxBuffer)

        window.showInformationMessage(
          `Exported: ${path.basename(saveUri.fsPath)}`,
        )
      } finally {
        try {
          await unlink(htmlTmpPath)
        } catch {
          // ignore
        }
      }
    },
  )
}
