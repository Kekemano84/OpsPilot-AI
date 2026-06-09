# OpsPilot AI

A full Flask SaaS MVP for self-employed finance, logistics yard checks and warehouse operations.

## Included

### Free
- Login / Register
- Plan selection
- Mileage
- Expenses
- Invoices
- PDF invoice download
- Tax estimate
- Yard Check demo, not saveable

### Pro - £4.99/month
- Yard Check Pro with saving
- Excel export
- Voice trailer logging using browser speech recognition
- Daily handover report
- KPI dashboard

### Business - £8.99/month
- Multi-user / team accounts
- Operations dashboard
- AI shift planner
- Photo trailer recognition workflow with photo upload, trailer/location/damage record and photo history

## Local setup

```bash
pip install -r requirements.txt
python app.py
```

Open:

```text
http://127.0.0.1:5000
```

## Render deployment

Build command:

```bash
pip install -r requirements.txt
```

Start command:

```bash
gunicorn app:app
```

## Demo note

This is a working MVP. Payments are simulated through plan selection. Real Stripe, real OCR and real AI APIs can be connected later.


## Added in Ultimate version

- Stripe-ready subscription routes
- Demo plan switching still available
- Invoice email sending via SMTP
- PDF handover download
- Team permissions
- Company/site fields
- Optional OpenAI-powered AI shift planner
- Optional OpenAI-powered photo recognition summary
- `.env.example` for deployment configuration

## Important setup notes

The app works without external APIs using fallback/demo mode.

For real payments:
1. Create Stripe products/prices.
2. Add `STRIPE_SECRET_KEY`, `STRIPE_PRO_PRICE_ID`, `STRIPE_BUSINESS_PRICE_ID`.
3. Set `APP_BASE_URL` to your Render URL.

For email invoice sending:
1. Add SMTP details.
2. Gmail usually needs an App Password.

For stronger AI:
1. Add `OPENAI_API_KEY`.
2. The app will use AI for shift planning and photo recognition summaries.


## Yard Check Batch Update

- Quick Add one trailer and remain on page
- Batch add multiple trailers at once
- Door 1-100 dropdown
- Fence 1-120 dropdown
- Custom locations
- Voice input maps Door/Fence values where possible
- Yard history shows latest 300 records

## Google Maps Mileage Calculation

Mileage page now supports postcode/address based driving mileage calculation.

Add this environment variable locally or on Render:

```text
GOOGLE_MAPS_API_KEY=your_google_maps_api_key
```

Required Google API:
- Distance Matrix API

Without the key, users can still enter miles manually.

## Editable Mileage Rate

- Mileage rate is editable per entry.
- Default rate can be changed in Settings.
- Each mileage record saves its own rate.

## Mileage Excel Export

Mileage History now has Download Excel:
- Date
- From
- To
- Purpose
- Miles
- Rate
- Claim
- Totals

## Expenses Excel Export

Expense History now has Download Excel:
- Date
- Category
- Description
- Amount
- Total

## Product focus update

Invoices and Tax Estimate were removed from the main navigation.

Free now includes:
- Mileage
- Expenses
- Yard Check demo only, not saved

Pro saves Yard Check records and exports Excel.

## Admin account

An admin/testing account is created automatically:

Email: admin@opspilot.ai
Password: admin123

Admin has Business access and can test all Pro and Business pages.
Admin Dashboard:
- Users
- Plan switching
- Role switching
- Global counts

## Yard settings and search

- Door count is editable in Settings.
- Fence count is editable in Settings.
- Yard Check page has search by trailer, door, fence, status, notes and source.
- Door/Fence dropdowns update from user settings.

## Stable fix

- Yard Check no longer crashes if older database is missing door_count/fence_count.
- Defaults: Door 1-100 and Fence 1-120.
- Settings can update Door Count and Fence Count.

## FLM Focus Version

Kept:
- Dashboard
- Morning Brief Builder
- Mileage
- Expenses
- Yard Check
- Handover
- Team
- Plans
- Admin only for admin user

Removed from navigation:
- KPI Dashboard
- Operations
- AI Planner
- Photo AI
- Business plan

Free:
- Dashboard
- Morning Brief
- Handover
- Team
- Mileage
- Expenses
- Yard Check demo only

Pro:
- Yard Check save/edit/delete
- Yard Check search/export
- Door/Fence settings


## Stripe Billing Added

Admin → Billing lets you set:
- Stripe publishable key
- Stripe secret key
- Webhook secret
- Pro monthly price ID
- App base URL
- Billing mode: rolling monthly or calendar-month prorated

Plans:
- Free
- Pro £4.99/month

Required Stripe setup:
1. Create Stripe account.
2. Create product: OpsPilot AI Pro.
3. Create recurring monthly price: £4.99.
4. Copy the Price ID into Admin → Billing.
5. Add webhook endpoint: /stripe/webhook.
6. In Stripe payout settings, add/update your bank account. Do not store card details in this app.


## App fixed version

Fixed:
- Removed KPI wording from dashboard/header.
- Added Yard Check edit/delete routes.
- Admin/Billing are admin-only in sidebar.
- Plans text is Free/Pro focused.
- Added PWA support: manifest.json, service-worker.js, app icons.
- App can be installed from Chrome/Edge as a standalone app.


## Admin Manual Pro / Trial Accounts

Admin can:
- Set a user to Free
- Set Manual Pro with no expiry
- Give 30 day Pro trial
- Give Gift Pro with no expiry

Trial Pro expires automatically and returns the user to Free.


## Camera and Bulk Voice

- Handover has Take Photo and Upload Photo.
- Yard Check has Take Photo and Upload Photo.
- Yard Check Batch Add has Bulk Voice Logging.
- Bulk voice can add multiple trailer records into the batch list at once.


## Pro Yard Configuration

Free users:
- Door Capacity fixed at 1-100
- Fence Capacity fixed at 1-120
- Cannot change yard numbers

Pro users:
- Can change Door Start / Door End
- Can change Fence Start / Fence End
- Yard Check dropdowns and capacity cards update automatically


## FINAL RENDER VERSION

This package is prepared for GitHub + Render deployment.

Admin login:
- Email: admin@opspilot.ai
- Password: admin123

Render:
- Build Command: pip install -r requirements.txt
- Start Command: gunicorn app:app

Local:
- python app.py
