"use strict";

/**
 * 範囲と、範囲に合う列の値の選び方。Power BI に触らない純粋な計算だけを置く（試験で直接たたく）。
 *
 * 範囲の端が null のときは開いている（下限なし・上限なし）。スライダーを端まで寄せたときも null にする。
 * 値を入れて閉じると、ほかのスライサーでデータが変わって最小・最大が動いたときに、端の外の値が黙って落ちるため。
 */

export interface Range {
    from: number | null;
    to: number | null;
}

export const OPEN_RANGE: Range = { from: null, to: null };

export function isOpen(range: Range): boolean {
    return range.from === null && range.to === null;
}

/** 範囲に入るか。メジャーが空（null）の値は、範囲を決めたら外す（標準のビジュアルのフィルターの「次の値より大きい」と同じ） */
export function inRange(value: number | null, range: Range): boolean {
    if (isOpen(range)) return true;
    if (value === null || !Number.isFinite(value)) return false;
    if (range.from !== null && value < range.from) return false;
    if (range.to !== null && value > range.to) return false;
    return true;
}

/** 保存された範囲を読む。数でないもの・NaN は開いた端にする。下限が上限を超えていたら入れ替える */
export function rangeOf(objects: { range?: { from?: unknown; to?: unknown } } | undefined): Range {
    const num = (v: unknown): number | null => (typeof v === "number" && Number.isFinite(v) ? v : null);
    const from = num(objects?.range?.from);
    const to = num(objects?.range?.to);
    if (from !== null && to !== null && from > to) return { from: to, to: from };
    return { from, to };
}

/**
 * 数の入力を読む。桁区切り・全角・空白を許す。空なら null（開いた端）。読めなければ undefined（元に戻す）。
 * 書式が % のメジャーは、見えている数（10.0%）で打つので 100 で割る（「%」を付けても付けなくても同じ）
 */
export function parseNumberInput(text: string, percent: boolean): number | null | undefined {
    const t = text.normalize("NFKC").replace(/[,\s]/g, "").replace(/%$/, "");
    if (t === "") return null;
    if (!/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) return undefined;
    const n = Number(t);
    if (!Number.isFinite(n)) return undefined;
    return percent ? n / 100 : n;
}

/** 書式の文字列が % か（引用符の中の % は数えない） */
export function isPercentFormat(format: string | undefined): boolean {
    if (!format) return false;
    return /%/.test(format.replace(/"[^"]*"|'[^']*'|\\./g, ""));
}
