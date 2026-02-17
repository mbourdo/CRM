# Annuity Funnel CRM

Simple browser-based CRM for Meta ads annuity funnel leads.

## Run

Preferred (persistent file storage):

1. Run `node server.js`
2. Open `http://localhost:3000`

If port `3000` is busy, run `PORT=4173 node server.js` and open `http://localhost:4173`.

This saves leads to `/Users/jake/Desktop/CRM/data/leads.json`.

Fallback mode:

- You can still open `index.html` directly, but that relies on browser localStorage only.

### Easiest Local Start (Mac)

1. Double-click `/Users/jake/Desktop/CRM/start-crm.command`
2. Your browser opens automatically
3. Keep that Terminal window open while using CRM

Data is saved to:

- `/Users/jake/Desktop/CRM/data/leads.json`

Important:

- For best data safety, always open via `http://localhost:3000` (or `4173`) from the launcher.
- Do not use `index.html` directly if you want strongest persistence.

## Host Online Free (Beginner Guide)

You said you already made a Supabase project. Great. Do these exact steps:

1. In Supabase, open `SQL Editor` and run this:

```sql
create table if not exists public.crm_state (
  id int primary key,
  leads jsonb not null default '[]'::jsonb,
  updated_at timestamptz not null default now()
);

insert into public.crm_state (id, leads)
values (1, '[]'::jsonb)
on conflict (id) do nothing;

alter table public.crm_state enable row level security;

drop policy if exists "public can read crm_state" on public.crm_state;
drop policy if exists "public can upsert crm_state" on public.crm_state;

create policy "public can read crm_state"
on public.crm_state
for select
to anon
using (true);

create policy "public can upsert crm_state"
on public.crm_state
for all
to anon
using (id = 1)
with check (id = 1);
```

2. In Supabase, go to `Settings` -> `API`.
3. Copy:
   - `Project URL`
   - `anon public` key
4. Open `/Users/jake/Desktop/CRM/config.js`.
5. Paste those two values:

```js
window.CRM_CONFIG = {
  supabaseUrl: "https://YOUR-PROJECT.supabase.co",
  supabaseAnonKey: "YOUR-ANON-KEY",
};
```

6. Test locally:
   - open `/Users/jake/Desktop/CRM/index.html`
   - create/edit one lead
   - refresh page and confirm it remains

7. Deploy frontend free (Netlify easiest):
   - go to [Netlify](https://netlify.com)
   - click `Add new site` -> `Deploy manually`
   - drag the whole `/Users/jake/Desktop/CRM` folder into Netlify
   - open your new Netlify URL

Your CRM will now use Supabase online data, and you can use it from anywhere.

## What it tracks

- Lead intake: first name, last name, email, phone
  - optional nickname (displayed as `First "Nickname" Last`)
- Appointment workflow:
  - separate 1st and 2nd appointment status
  - separate 1st and 2nd confirmations (phone/calendar)
  - separate 1st and 2nd appointment time
- Call activity:
  - one-click event logging from lead card:
    - Call Live Contact / Call VM / Call NC
    - Text Sent / Text Received
  - automatic timestamp on each log event
  - per-log note field with save action
- Sales process:
  - first appointment status
  - annuity presentation (2nd) status
  - documents status
  - carrier outreach status
- Follow-up management:
  - `Next Action` text
  - date/time follow-up reminder
- Stage SLA highlighting:
  - overdue leads are visually flagged based on stage age
- Lead health statuses:
  - active / no show / inactive / rescheduled / do not contact
  - quick actions on card for no show, inactive, and set active
- Opportunity size:
  - treated as estimated annuity premium (`Est. Premium`)
  - accepts currency-style input like `$263,000`
- Close management:
  - won/lost stages
  - close reason
  - close date
- Notes
- Pipeline stages:
  - new lead -> booked -> first meeting -> annuity presentation -> carrier outreach -> proposal -> won/lost
- KPIs:
  - total leads, booked, booked unconfirmed, phone/calendar confirmed, live contacts, open pipeline premium, closed won

## CSV

- Export all leads to CSV (`Export CSV`)
- Import leads from CSV (`Import CSV`)
- Import does ID-based upsert: same `id` updates, missing/new `id` inserts

## Previous version notes

- Existing stored leads are auto-normalized to new fields on load.

## Legacy fields

- The app still keeps these booleans for process tracking:
  - confirmed by phone
  - confirmed in calendar

## Data storage

Primary: `/Users/jake/Desktop/CRM/data/leads.json` (when running `server.js`).

Online mode: Supabase table `public.crm_state` (when `config.js` has Supabase keys).

Fallback: browser localStorage (`annuity-crm-leads-v1`) if API is unavailable.
