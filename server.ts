import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import apiApp from "./artifacts/api-server/src/app";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const port = 3000;

  // Mount API routes
  app.use(apiApp);

  // In development, mount Vite dev server with HMR and React fast refresh
  if (process.env.NODE_ENV !== "production") {
    const { createServer: createViteServer } = await import("vite");
    const vite = await createViteServer({
      server: { middlewareMode: true, host: "0.0.0.0" },
      appType: "spa",
      root: path.resolve(__dirname, "artifacts/maquina-accounts"),
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.resolve(__dirname, "artifacts/maquina-accounts/dist/public");
    app.use(express.static(distPath));
    app.get("*", (_req, res) => {
      res.sendFile(path.resolve(distPath, "index.html"));
    });
  }

  app.listen(port, "0.0.0.0", () => {
    console.log(`Server listening on http://0.0.0.0:${port}`);
  });
}

startServer().catch((err) => {
  console.error("Failed to start server:", err);
  process.exit(1);
});
