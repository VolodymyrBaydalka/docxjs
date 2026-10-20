// Bullets in the Symbol and Wingdings fonts use private-use code points (U+F0xx)
// that only those fonts map to a glyph. Without the font (macOS has no Symbol
// mapping for them, Linux and most browsers have neither font) the bullet shows
// as an empty box. Map the bullets Word uses by default to Unicode and drop the font.
const symbolFontBullets: Record<string, Record<string, string>> = {
	symbol: { "\uf0b7": "\u2022" }, // •
	wingdings: {
		"\uf06c": "\u25cf", // ●
		"\uf06e": "\u25a0", // ■
		"\uf071": "\u2751", // ❑
		"\uf075": "\u25c6", // ◆
		"\uf076": "\u2756", // ❖
		"\uf0a7": "\u25aa", // ▪
		"\uf0d8": "\u27a2", // ➢
		"\uf0e8": "\u2794", // ➔
		"\uf0fb": "\u2717", // ✗
		"\uf0fc": "\u2713", // ✓
	},
};

export function symbolFontBullet(text: string, style: Record<string, string>): { text: string, style: Record<string, string> } {
	const font = (style?.["font-family"] ?? "").split(",")[0].replace(/['"]/g, "").trim().toLowerCase();
	const map = symbolFontBullets[font];

	if (!map || !text || ![...text].some(ch => map[ch]))
		return { text, style };

	const { "font-family": _, ...rest } = style;
	return { text: [...text].map(ch => map[ch] ?? ch).join(""), style: rest };
}
