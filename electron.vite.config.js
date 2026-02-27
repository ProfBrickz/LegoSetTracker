// Imports
import { defineConfig } from "electron-vite";
import path from "path";
import { normalizePath } from "vite";
import { viteStaticCopy } from "vite-plugin-static-copy";


export default defineConfig({
	main: {
		build: {
			outDir: path.resolve(__dirname, "out/main"),
			watch: {},
			rollupOptions: {
				input: {
					main: path.resolve(__dirname, "src/main.js")
				}
			}
		},
	},
	preload: {
		build: {
			watch: {},
			rollupOptions: {
				input: {
					preload: path.resolve(__dirname, "src/preload/preload.js")
				},
				output: {
					format: "cjs",
					entryFileNames: "preload.js"
				}
			}
		}
	},
	renderer: {
		root: path.resolve(__dirname, "src/renderer"),
		build: {
			outDir: path.resolve(__dirname, "out/renderer"),
			watch: {},
			rollupOptions: {
				input: path.resolve(__dirname, "src/renderer/views/layouts/main.html")
			}
		},
		plugins: [
			viteStaticCopy({
				targets: [
					{
						src: normalizePath(path.resolve(__dirname, "src/renderer/views/pages")),
						dest: "views"
					}
				],
			})
		]
	}
});
