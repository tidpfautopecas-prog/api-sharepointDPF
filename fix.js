const accessToken = '<SEU_TOKEN_AQUI>';

const res = await fetch(
  'https://graph.microsoft.com/v1.0/sites?search=DPF-FS',
  { headers: { 'Authorization': `Bearer ${accessToken}` } }
);

const data = await res.json();
console.log(data);
