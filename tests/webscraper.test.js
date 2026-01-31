// Imports
import { jest } from "@jest/globals";
import { getWebpage, getColors } from "../src/webscraper.js";


// Tests
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

describe("getColors", () => {
	beforeEach(() => {
		global.fetch = jest.fn();
	});

	afterEach(() => {
		jest.resetAllMocks();
	});

	test("returns an array of colors", async () => {
		const html = `
			<!doctype html>
			<html>
				<body>
					<div>
						<div>
							<table></table>
							<table>
								<tbody>
									<tr>
										<td></td>
										<td>
											<p>White</p>
											<span>
												LEGO
												Color: <!-- -->White - 1
											</span>
										</td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td>1</td>
									</tr>
									<tr>
										<td></td>
										<td>
											<p>Red</p>
											<span>
												LEGO
												Color: <!-- -->Bright Red - 21
											</span>
										</td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td>5</td>
									</tr>
									<tr>
										<td></td>
										<td>
											<p>Sand Green</p>
											<span>
												LEGO
												Color: <!-- -->Sand Green - 151
											</span>
										</td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td>48</td>
									</tr>
								</tbody>
							</table>
						</div>
						<div>
							<table></table>
							<table>
								<tbody>
									<tr>
										<td></td>
										<td>
											<p>Trans-Clear</p>
											<span>
												LEGO
												Color: <!-- -->Transparent - 40
											</span>
										</td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td>12</td>
									</tr>
									<tr>
										<td></td>
										<td>
											<p>Trans-Red</p>
											<span>
												LEGO
												Color: <!-- -->Tr. Red - 41
											</span>
										</td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td>17</td>
									</tr>
									<tr>
										<td></td>
										<td>
											<p>Trans-Dark Blue</p>
											<span>
												LEGO
												Color: <!-- -->Tr. Blue - 43
											</span>
										</td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td>14</td>
									</tr>
								</tbody>
							</table>
						</div>
						<div>
							<table></table>
							<table>
								<tbody>
									<tr>
										<td></td>
										<td>
											<p>Chrome Black</p>
											<span>LEGO Color: </span>
										</td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td>122</td>
									</tr>
								</tbody>
							</table>
						</div>
					</div>
				</body>
			</html>
		`;

		fetch.mockResolvedValue({
			ok: true,
			text: jest.fn().mockResolvedValue(html)
		});

		const colors = await getColors();

		expect(colors).toBeInstanceOf(Array);
		expect(colors.length).toEqual(7);
		expect(colors).toEqual([
			{ bricklinkId: 1, bricklinkName: "White", legoId: 1, legoName: "White" },
			{ bricklinkId: 5, bricklinkName: "Red", legoId: 21, legoName: "Bright Red" },
			{ bricklinkId: 48, bricklinkName: "Sand Green", legoId: 151, legoName: "Sand Green" },
			{ bricklinkId: 12, bricklinkName: "Trans-Clear", legoId: 40, legoName: "Transparent" },
			{ bricklinkId: 17, bricklinkName: "Trans-Red", legoId: 41, legoName: "Tr. Red" },
			{ bricklinkId: 14, bricklinkName: "Trans-Dark Blue", legoId: 43, legoName: "Tr. Blue" },
			{ bricklinkId: 122, bricklinkName: "Chrome Black", legoId: null, legoName: null },
		]);
	});

	test("returns an empty array", async () => {
		const html = `
			<!doctype html>
			<html>
				<body>
				</body>
			</html>
		`;

		fetch.mockResolvedValue({
			ok: true,
			text: jest.fn().mockResolvedValue(html)
		});

		const colors = await getColors();

		expect(colors).toBeInstanceOf(Array);
		expect(colors.length).toEqual(0);
		expect(colors).toEqual([]);
	});

	test("throws an error", async () => {
		fetch.mockResolvedValue({
			ok: false,
			status: 404,
			statusText: "Not Found"
		});

		await expect(getColors()).rejects.toThrow("Failed to fetch https://v2.bricklink.com/en-us/catalog/color-guide: 404 Not Found");
	});
});
