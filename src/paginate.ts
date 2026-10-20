// Splits rendered pages whose content is taller than the page, the way Word
// lays out text: whole blocks move to a new page, paragraphs split at a line
// boundary (at least two lines per page), tables split between rows with their
// header rows (<thead>) repeated, and a block marked keep-with-next
// (break-after: avoid) moves along with the block after it.
//
// It measures the layout, so the pages must be in the document and the fonts
// loaded. renderAsync runs it with the `paginate` option; call it yourself
// after renderDocument when you insert the nodes elsewhere.

const MAX_PAGES = 2000;

/** Splits every rendered page (section.<className>) under root until each fits its page height. */
export function paginate(root: HTMLElement, className: string = "docx"): void {
	const pages = Array.from(root.querySelectorAll<HTMLElement>(`section.${className}`));

	for (let i = 0; i < pages.length && pages.length < MAX_PAGES; i++) {
		const next = splitPage(pages[i], className);

		if (next)
			pages.splice(i + 1, 0, next);
	}
}

/**
 * Moves the blocks that do not fit on a page to a new page right after it.
 * Returns the new page, or null when the content fits or cannot be split.
 */
export function splitPage(page: HTMLElement, className: string = "docx"): HTMLElement | null {
	const article = Array.from(page.children).find(c => c.tagName === "ARTICLE") as HTMLElement;

	if (!article)
		return null;

	const style = getComputedStyle(page);
	const pageHeight = parseFloat(style.minHeight);

	if (!pageHeight)
		return null;

	let available = pageHeight - (parseFloat(style.paddingTop) || 0) - (parseFloat(style.paddingBottom) || 0);

	for (const c of Array.from(page.children)) {
		if (c !== article)
			available -= c.getBoundingClientRect().height;  // header and footer
	}

	const blocks = Array.from(article.children) as HTMLElement[];
	const limitY = article.getBoundingClientRect().top + available + 0.5;
	// Empty paragraphs (such as the one holding a page break) never start a page.
	const over = blocks.findIndex(b => !isBlank(b) && b.getBoundingClientRect().bottom > limitY);

	if (over < 0) {
		for (const b of blocks) {
			if (isBlank(b) && b.getBoundingClientRect().bottom > limitY)
				b.style.display = "none";
		}

		return null;
	}

	let cut = over;
	let carry: HTMLElement = null;

	if (blocks[over].tagName === "TABLE")
		carry = splitTable(blocks[over] as HTMLTableElement, limitY);
	else if (blocks[over].tagName === "P")
		carry = splitParagraph(blocks[over], limitY, className);

	if (carry) {
		cut = over + 1;
	} else {
		while (cut > 0 && getComputedStyle(blocks[cut - 1]).breakAfter === "avoid")
			cut--;

		if (cut === 0)
			cut = over || 1;

		if (cut >= blocks.length)
			return null;  // a single block taller than a page stays where it is
	}

	const next = page.cloneNode(false) as HTMLElement;
	let nextArticle: HTMLElement = null;

	for (const c of Array.from(page.children)) {
		if (c === article) {
			nextArticle = article.cloneNode(false) as HTMLElement;
			next.appendChild(nextArticle);
		} else {
			next.appendChild(c.cloneNode(true));
		}
	}

	if (carry)
		nextArticle.appendChild(carry);

	for (const b of blocks.slice(cut))
		nextArticle.appendChild(b);

	page.after(next);
	return next;
}

/**
 * Splits a table at the first body row that ends below limitY. The rows from
 * there move to a new table with the same attributes, columns and header rows.
 * Returns that table, or null when no body row fits above limitY.
 */
export function splitTable(table: HTMLTableElement, limitY: number): HTMLTableElement | null {
	const head = table.tHead;
	const rows = Array.from(table.rows).filter(r => r.parentElement !== head);
	const at = rows.findIndex(r => r.getBoundingClientRect().bottom > limitY);

	if (at <= 0)
		return null;

	const rest = table.cloneNode(false) as HTMLTableElement;

	for (const c of Array.from(table.children)) {
		if (c.tagName === "COLGROUP" || c === head)
			rest.appendChild(c.cloneNode(true));
	}

	for (const r of rows.slice(at))
		rest.appendChild(r);

	return rest;
}

/**
 * Splits a paragraph at a line boundary so the lines ending above limitY stay,
 * with at least two lines on each side (Word's widow and orphan control).
 * Returns the continuation (inserted after p), or null when it cannot split.
 */
export function splitParagraph(p: HTMLElement, limitY: number, className: string = "docx"): HTMLElement | null {
	if (getComputedStyle(p).breakInside === "avoid")
		return null;

	const lines = paragraphLines(p);
	let keep = lines.filter(l => l.bottom <= limitY).length;

	if (lines.length - keep < 2)
		keep = lines.length - 2;

	if (keep < 2)
		return null;

	const at = textPositionAt(p, lines[keep].top - 1);

	if (!at)
		return null;

	const range = document.createRange();
	range.setStart(at.node, at.offset);
	range.setEnd(p, p.childNodes.length);

	const rest = p.cloneNode(false) as HTMLElement;
	rest.removeAttribute("id");
	rest.appendChild(range.extractContents());

	// The first part ends the page: no space after it. The continuation is not
	// a new paragraph: no space before it, no first-line indent, no list marker.
	p.style.marginBottom = "0";
	rest.style.marginTop = "0";
	rest.style.textIndent = "0";
	rest.classList.add(`${className}-continued`);

	p.after(rest);
	return rest;
}

/** Line boxes of a paragraph ({top, bottom}, in order), from its text rectangles. */
export function paragraphLines(p: HTMLElement): { top: number, bottom: number }[] {
	const range = document.createRange();
	range.selectNodeContents(p);

	const lines: { top: number, bottom: number }[] = [];

	for (const r of Array.from(range.getClientRects())) {
		if (!r.width && !r.height)
			continue;

		const line = lines.find(l => Math.abs(l.top - r.top) < 2 || (r.top < l.bottom - 2 && r.bottom > l.top + 2));

		if (line) {
			line.top = Math.min(line.top, r.top);
			line.bottom = Math.max(line.bottom, r.bottom);
		} else {
			lines.push({ top: r.top, bottom: r.bottom });
		}
	}

	return lines.sort((a, b) => a.top - b.top);
}

/** Text position of the first character on or below y, by binary search over the text nodes. */
function textPositionAt(p: HTMLElement, y: number): { node: Text, offset: number } | null {
	const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
	const range = document.createRange();

	const topOf = (node: Text, offset: number) => {
		range.setStart(node, offset);
		range.setEnd(node, offset + 1);
		const rects = range.getClientRects();
		return rects.length ? rects[0].top : -Infinity;
	};

	for (let n = walker.nextNode() as Text; n; n = walker.nextNode() as Text) {
		if (!n.length || topOf(n, n.length - 1) < y)
			continue;

		let lo = 0;
		let hi = n.length - 1;

		while (lo < hi) {
			const mid = (lo + hi) >> 1;

			if (topOf(n, mid) >= y)
				hi = mid;
			else
				lo = mid + 1;
		}

		return { node: n, offset: lo };
	}

	return null;
}

function isBlank(el: HTMLElement): boolean {
	return !el.textContent.trim() && !el.querySelector("img, svg, table, canvas, image");
}
