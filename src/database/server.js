// Imports
import { PGlite } from "@electric-sql/pglite";
import { PGLiteSocketServer } from '@electric-sql/pglite-socket';
import path from "path";


// Constants
const databasePort = 5432;


// Create a PGlite instance
const database = new PGlite(path.join(import.meta.dirname, "../../database"));


// Create and start a socket server
const server = new PGLiteSocketServer({
	db: database,
	port: databasePort,
	host: '127.0.0.1',
	inspect: true,
	debug: true
});

await server.start();
console.log(`Server started on 127.0.0.1:${databasePort}`);

// Handle graceful shutdown
process.on('SIGINT', async () => {
	await server.stop();
	await database.close();
	console.log('Server stopped and database closed');
	process.exit(0);
});
