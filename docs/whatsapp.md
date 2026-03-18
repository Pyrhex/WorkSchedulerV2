# WhatsApp Image Sender (Twilio Sandbox)

The schedule view includes a **Send WhatsApp Image** button that lets schedulers paste or drop a screenshot and push it straight to the Twilio WhatsApp Sandbox number for free testing.

## Prerequisites
1. Enable the WhatsApp Sandbox in the Twilio Console and note the join code.
2. On the handset that should receive messages, send `join <your-code>` to the sandbox number `+1 415 523 8886` via WhatsApp so it becomes a permitted recipient.
3. Expose the app (or at least the static uploads folder) over HTTPS so Twilio can download pasted images. The images are stored in `static/whatsapp-media`, so the publicly reachable URL must map to that folder.

## Environment variables
Set the following in `.env` (or your deployment environment):

| Variable | Description |
| --- | --- |
| `TWILIO_ACCOUNT_SID` | Your Twilio Account SID. |
| `TWILIO_AUTH_TOKEN` | The matching Auth Token (use a test credential for sandboxing). |
| `TWILIO_WHATSAPP_FROM` | (Optional) WhatsApp sender value. Defaults to Twilio's sandbox ID `whatsapp:+14155238886`. |
| `TWILIO_WHATSAPP_TO` | Destination WhatsApp number (must include the `whatsapp:` prefix or it will be added automatically). This number has to opt-in to the sandbox. |
| `WHATSAPP_MEDIA_BASE_URL` | Public base URL that maps to `static/whatsapp-media` (e.g., `https://example.com/static/whatsapp-media`). Twilio downloads images from this URL. |

## Usage flow
1. Open the schedule page and click **Send WhatsApp Image**.
2. Paste (`Ctrl/Cmd+V`) or drop an image, optionally add a caption, and hit **Send to WhatsApp**.
3. The backend saves the image under `static/whatsapp-media`, builds a public URL using `WHATSAPP_MEDIA_BASE_URL`, and posts it to Twilio's `Messages` API using your sandbox credentials.
4. Twilio forwards the media message from the sandbox number to the opted-in handset.

Uploads older than 24 hours are automatically pruned from `static/whatsapp-media` every time a new message is sent.
