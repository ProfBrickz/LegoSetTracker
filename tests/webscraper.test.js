import { jest } from "@jest/globals";
import { JSDOM } from "jsdom";
import { getWebpage } from "../src/webscraper.js";


describe("getWebpage", () => {
	beforeEach(() => {
		global.fetch = jest.fn();
	});

	afterEach(() => {
		jest.resetAllMocks();
	});

	test("fetches the URL and returns a document on success", async () => {
		const html = `
			<!doctype html>
			<html>
				<body>
					<h1>Hello World!</h1>
				</body>
			</html>
		`;

		fetch.mockResolvedValue({
			ok: true,
			text: jest.fn().mockResolvedValue(html),
		});

		const document = await getWebpage("https://example.com");

		expect(fetch).toHaveBeenCalledWith("https://example.com");
		expect(document.constructor.name).toEqual("Document");
		expect(document.documentElement.tagName).toEqual("HTML");
		expect(document.querySelector("h1").textContent).toBe("Hello World!");
	});

	test("throws error when the response is not ok", async () => {
		fetch.mockResolvedValue({
			ok: false,
			status: 404,
			statusText: "Not Found"
		});

		await expect(getWebpage("https://example.com")).rejects.toThrow(
			new Error("Failed to fetch https://example.com: 404 Not Found")
		);
	});
});
