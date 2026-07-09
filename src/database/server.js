// Imports
import { PGlite } from "@electric-sql/pglite";
import { PGLiteSocketServer } from '@electric-sql/pglite-socket';
import path from "path";


// Constants
const DATABASE_HOST = '127.0.0.1';;
const DATABASE_PORT = 5432;


// Create a PGlite instance
const database = new PGlite(path.join(import.meta.dirname, "../../data/database"));


// Create and start a socket server
const server = new PGLiteSocketServer({
	db: database,
	host: DATABASE_HOST,
	port: DATABASE_PORT,
	inspect: true,
	debug: true,
	maxConnections: 10
});

await server.start();
console.log(`Server started on ${DATABASE_HOST}:${DATABASE_PORT}`);

// Handle graceful shutdown
process.on('SIGINT', async () => {
	await server.stop();
	await database.close();
	console.log('Server stopped and database closed');
	process.exit(0);
});
