# Viber vs Doomer PowerPoint Add-in

This repository implements the **Agentic Engineering "Viber vs Doomer"** feedback integration as a Microsoft PowerPoint web add-in. It enables presenters to insert live survey result placeholders in slides, keep the companion mobile application synchronized with slide navigation, and display real-time audience feedback directly inside the presentation.

## Features

- **Ribbon integration** with commands to insert live result placeholders, open the configuration panel, and refresh every embedded instance.
- **Task pane control center** to manage the presentation ID, map slides to survey identifiers, and fetch presentation metadata from the Agentic Feedback API.
- **Content add-in** that renders live survey charts while in presentation mode and shows informative placeholders in edit mode.
- **Automatic slide synchronization** that notifies the API whenever the presenter changes slides, throttled to one request every two seconds.
- **Manual refresh broadcast** using document storage tokens to force all open instances to reload.

## Project Structure

```
public/
  commands.html      # Ribbon command surface loader
  commands.js        # Ribbon handlers and slide sync logic
  common.js          # Shared storage, PowerPoint, and API utilities
  content.html       # Content add-in host page for live charts
  content.js         # Live polling, chart rendering, and placeholders
  styles.css         # Shared styling across task pane and content
  taskpane.html      # Configuration panel host page
  taskpane.js        # Presentation/slide mapping UX logic
manifest.xml          # Office add-in manifest referencing the hosted assets
server.js             # Lightweight static HTTPS/HTTP server for local dev
package.json          # NPM metadata with convenience scripts
```

## Getting Started

1. **Install dependencies** (Node.js 18+ required). There are no external packages to install for the runtime scripts, but you can run `npm install` to capture any future tooling additions.
2. **Generate HTTPS certificates** for local development (required by Office add-ins). You can use the [Office Add-in certificate utility](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#enable-https).
3. **Start the static server**:

   ```bash
   SSL_KEY_PATH=path/to/server.key \
   SSL_CERT_PATH=path/to/server.crt \
   npm start
   ```

   The add-in will be hosted at `https://localhost:3000/`.

4. **Sideload the add-in** in PowerPoint by uploading `manifest.xml` via `Insert > My Add-ins > Upload My Add-in`.
5. **Use the ribbon commands**:
   - *Insert Live Results* adds a placeholder textbox and saves default mappings.
   - *Configure Slide Mapping* opens the task pane for granular control.
   - *Refresh All Data* bumps the refresh token so every live results instance pulls the newest data.

6. **Present and collect feedback**. When you enter slideshow mode, each live results content add-in polls `https://api.agenticfeedback.app` every two seconds, rendering a bar chart of responses. Slide transitions automatically notify the API so the audience app stays in sync.

## Configuration Notes

- The default presentation ID is `boston-code-camp-2025`. Update it from the task pane if you need to point at another presentation.
- Each slide can override both the `slide_id` (used for slide sync) and an optional `survey_id` (used for result polling) via the control panel.
- Document settings persist inside the PowerPoint file, so mappings travel with the deck.

## API Usage

All network calls leverage the documented endpoints:

- `POST https://api.agenticfeedback.app/slides/current` to update the currently visible slide for the audience.
- `GET https://api.agenticfeedback.app/slides/{slide_id}/results` for live aggregated responses.
- `GET https://api.agenticfeedback.app/presentations/{presentation_id}` to fetch metadata for configuration assistance.

Requests are throttled to at most one update every two seconds to prevent flooding the API.

## Limitations & Future Enhancements

- The ribbon command inserts a styled textbox as a placeholder; presenters should replace it with the content add-in (`Insert > My Add-ins > Viber vs Doomer Live Feedback`) if the automatic insertion is insufficient on certain platforms.
- Rendering leverages simple CSS-based bar charts instead of Chart.js to avoid external dependencies in constrained environments. This can be swapped with a richer library when available.
- Authentication is not implemented; if the API introduces protected endpoints, integrate secure storage via `OfficeRuntime.storage` or similar facilities.

## Scripts

- `npm start` – Launches the static development server (HTTPS when certificate paths are supplied).
- `npm run start:http` – Same as `npm start`; provided as a semantic alias for HTTP-only debugging.

## License

Distributed under the MIT license. See [LICENSE](LICENSE).
