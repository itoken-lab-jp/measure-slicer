"use strict";

/**
 * 絞る条件。範囲（メジャーの値が下限〜上限）、上位 N・下位 N（メジャーの順位）のどれかと、除外（条件に合うものを除く）。
 * Power BI に触らない純粋な計算だけを置く。
 */

import { Range, inRange, isOpen } from "./range";
import { RowValue } from "./filter";

export const MODES = { range: "range", top: "top", bottom: "bottom" } as const;
export type Mode = (typeof MODES)[keyof typeof MODES];

export interface Condition {
    mode: Mode;
    range: Range;
    /** 上位・下位の件数。null は未指定（絞らない） */
    count: number | null;
    exclude: boolean;
}

/** 条件を決めていないか（範囲が開いている、件数が未指定）。決めていなければ除外でも絞らない */
export function isUnset(condition: Condition): boolean {
    return condition.mode === MODES.range ? isOpen(condition.range) : condition.count === null;
}

/**
 * 行ごとに、絞り込みを通るか。
 * 上位・下位は空の値を除いて順位を付け、N 番目と同じ値も入れる（DAX の TOPN と同じく、同順位で切らない）。
 * 除外は、条件に合う行を落として残りを通す（空の値の行も残る）
 */
export function keptRows(rows: RowValue[], condition: Condition): boolean[] {
    if (isUnset(condition)) return rows.map(() => true);
    let matched: boolean[];
    if (condition.mode === MODES.range) {
        matched = rows.map((r) => inRange(r.value, condition.range));
    } else {
        const values = rows.map((r) => r.value).filter((v): v is number => v !== null);
        const sorted = [...values].sort((a, b) => (condition.mode === MODES.top ? b - a : a - b));
        const n = Math.max(0, Math.floor(condition.count ?? 0));
        const edge = n === 0 ? null : sorted[Math.min(n, sorted.length) - 1] ?? null;
        matched = rows.map((r) => {
            if (r.value === null || edge === null) return false;
            return condition.mode === MODES.top ? r.value >= edge : r.value <= edge;
        });
    }
    return condition.exclude ? matched.map((m) => !m) : matched;
}

/** 保存された件数を読む。1 以上の整数だけ（それ以外は未指定） */
export function countOf(value: unknown): number | null {
    return typeof value === "number" && Number.isFinite(value) && value >= 1 ? Math.floor(value) : null;
}

export function modeOf(value: unknown): Mode {
    return value === MODES.top || value === MODES.bottom ? value : MODES.range;
}
