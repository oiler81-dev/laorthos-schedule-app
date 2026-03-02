module.exports = async function (context, req) {
  const headers = req.headers || {};

  const email =
    headers["x-ms-client-principal-name"] ||
    headers["X-MS-CLIENT-PRINCIPAL-NAME"] ||
    null;

  const editors = new Set([
    "nperez@unitymsk.com",
    "aledezma@laorthos.com"
  ]);

  const isEditor = email ? editors.has(email.toLowerCase()) : false;

  context.res = {
    status: 200,
    headers: { "Content-Type": "application/json" },
    body: {
      email,
      editor: isEditor
    }
  };
};
