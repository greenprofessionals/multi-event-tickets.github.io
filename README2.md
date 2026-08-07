# Multi-Event Ticketing System

This folder is a NEW standalone version. It does not replace or modify the existing working `/slpp-tickets/` system.

## Folder structure

- `index.html` — public event selector
- `claim.html` — public self-service claim page
- `v.html` — distributor batch + individual voucher page
- `admin.html` — multi-event control center
- `gate.html` — fast event-day check-in mode using the working QR scanner pattern
- `config.html` — event manager + logo/background upload
- `styles.css` — shared UI styling
- `shared.js` — API helpers + ticket renderer
- `site-config.js` — ONLY place to paste the new Apps Script endpoint and optional Google Form URL
- `Code.gs` — Google Apps Script backend for the NEW spreadsheet
- `images/claim-og.jpg`
- `images/config-og.jpg`
- `images/admin-og.jpg`
- `images/gate-og.jpg`

## New Google Sheet / Apps Script setup

1. Create a **new Google Sheet** for the multi-event system.
2. Open **Extensions → Apps Script**.
3. Replace starter code with `Code.gs` from this folder.
4. Edit `ADMIN_DIRECTORY` near the top of `Code.gs` before production use.
5. Run `setupMultiEventSystem()` once and authorize it.
6. Run `createEventConfigForm()` once. This creates the new Google Form and connects it to the new spreadsheet.
7. Deploy Apps Script as a Web App:
   - Execute as: **Me**
   - Who has access: **Anyone**
8. Copy the `/exec` Web App URL into `site-config.js`.
9. Copy the published Google Form URL into `FORM_URL` in `site-config.js` if you want the button on `config.html`.
10. Upload this entire `multi-event-tickets/` folder into the GitHub Pages repository.

Expected public path:

`https://greenprofessionals.github.io/multi-event-tickets/`

## Branding model

Branding is stored per Event ID. Every event can have its own:

- organization name
- chapter/unit name
- event title and tagline
- date/time/venue
- logo
- optional background image
- primary/accent colors
- ticket serial prefix
- currency symbol
- legal/footer text
- ticket tiers and capacities
- groups/chapters/tables
- overall venue capacity

The logo uploaded from `config.html` is automatically used as both the ticket's top-left logo and its large faint watermark.

## Multi-event data model

The new spreadsheet creates these sheets:

- Events
- Tiers
- Groups
- Vouchers
- Claims
- CheckIns
- Payments
- AuditLog
- Counters
- EventConfigResponses (created when the Google Form is created)

Every operational record contains an Event ID so events do not share counters, branding, reporting, tickets or capacity.

## Main V2 features

- Multiple events in one backend
- Per-event branding
- Public event selector
- Distributor voucher batches
- Capacity controls (event and ticket-tier level)
- Admin dashboard
- Guest search by name/phone/email/serial
- QR and manual check-in
- Walk-in sale/check-in
- Payment status, method and amount tracking
- Ticket transfer/correction
- Ticket revoke/reactivate
- QR credential reissue
- Audit log
- Event close/archive controls
- Attendance/revenue summary and check-ins by hour

## Important deployment note

This package cannot create the actual Google Sheet or Google Form from outside your Google account. The included Apps Script creates the required sheets and the connected Form when you run the two setup functions inside the NEW Sheet.

## QR scanner

`gate.html` intentionally preserves the scanner approach that is working on the current iPhone deployment: rear-camera request, `BarcodeDetector` when supported, jsQR fallback, photo scanning fallback, and automatic check-in after a successful read.
