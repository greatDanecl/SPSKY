export default async (request, context) => {
  const auth = request.headers.get("Authorization");
  
  // Define aquí tu contraseña (o usa variables de entorno)
  const password = "TU_CONTRASEÑA_AQUÍ";
  const expectedAuth = `Basic ${btoa(`usuario:${password}`)}`;

  if (auth !== expectedAuth) {
    return new Response("Acceso Denegado", {
      status: 401,
      headers: { "WWW-Authenticate": 'Basic realm="Acceso Protegido"' },
    });
  }

  return context.next();
};
