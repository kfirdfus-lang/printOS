# send-email Edge Function

This function sends emails through Resend.

## Requirements

- Supabase project linked locally (`supabase link`)
- Secret exists in Supabase:
  - `RESEND_API_KEY`

Set secret manually if needed:

```bash
supabase secrets set RESEND_API_KEY=your_resend_api_key
```

## Deploy function

```bash
supabase functions deploy send-email
```

## Invoke manually

From your project root:

```bash
supabase functions invoke send-email --no-verify-jwt --data '{"to":"test@example.com","subject":"Test","text":"Hello from send-email"}'
```

You can also send HTML:

```bash
supabase functions invoke send-email --no-verify-jwt --data '{"to":"test@example.com","subject":"HTML Test","html":"<h1>Hello</h1><p>From PrintOS</p>"}'
```
