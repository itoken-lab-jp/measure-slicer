"use strict";

/**
 * スライダーの上に描く分布。列の値ごとのメジャーの値が、最小〜最大のどこに何件あるか。
 * 標準の数値の範囲スライサーには無く、範囲を動かす前に「どこに製品が固まっているか」「どの製品が外れたか」を見せる。
 *
 * 件数が少なければ 1 件 1 点（乗せると名前と値）、多ければ値の幅で区切った棒にする。
 * 通るかどうかは keptRows の結果をそのまま使う（範囲だけでなく上位 N・下位 N・除外でも同じ見え方になる）。
 * Power BI に触らない純粋な計算だけを置く。
 */

import { RowValue } from "./filter";
import { maxOf } from "./shared/minMax";

export const SHAPES = { auto: "auto", bars: "bars", points: "points" } as const;
export type Shape = (typeof SHAPES)[keyof typeof SHAPES];

/** 自動のとき、点で描く件数の上限。これを超えたら棒 */
export const POINT_LIMIT = 60;
/** 棒の数 */
export const BIN_COUNT = 20;
/** 点を重ねる列の数。同じ列に入った点は上へ積む */
export const POINT_COLUMNS = 40;
/** 積める点の段の数。超える列があれば、自動では棒にする */
export const POINT_STACK_LIMIT = 4;

export interface DistributionPoint {
    /** 0〜1 の位置（最小が 0、最大が 1） */
    position: number;
    /** 下から何段目か（0 から） */
    stack: number;
    row: number;
    kept: boolean;
}

export interface DistributionBar {
    /** 0〜1 の左端 */
    start: number;
    count: number;
    kept: number;
}

export type Distribution =
    | { kind: "none" }
    | { kind: "points"; points: DistributionPoint[] }
    | { kind: "bars"; bars: DistributionBar[]; maxCount: number };

export function shapeOf(value: unknown): Shape {
    return value === SHAPES.bars || value === SHAPES.points ? value : SHAPES.auto;
}

function positionOf(value: number, min: number, max: number): number {
    return max === min ? 0.5 : Math.min(1, Math.max(0, (value - min) / (max - min)));
}

/** 値のある行の点。同じ列に入った点は上へ積む */
function pointsOf(rows: RowValue[], kept: boolean[], min: number, max: number): DistributionPoint[] {
    const used = new Map<number, number>();
    const points: DistributionPoint[] = [];
    rows.forEach((r, i) => {
        if (r.value === null) return;
        const position = positionOf(r.value, min, max);
        const column = Math.round(position * POINT_COLUMNS);
        const stack = used.get(column) ?? 0;
        used.set(column, stack + 1);
        points.push({ position, stack, row: i, kept: kept[i] });
    });
    return points;
}

function barsOf(rows: RowValue[], kept: boolean[], min: number, max: number): DistributionBar[] {
    const bars: DistributionBar[] = Array.from({ length: BIN_COUNT }, (_, i) => ({ start: i / BIN_COUNT, count: 0, kept: 0 }));
    rows.forEach((r, i) => {
        if (r.value === null) return;
        const bin = Math.min(BIN_COUNT - 1, Math.floor(positionOf(r.value, min, max) * BIN_COUNT));
        bars[bin].count++;
        if (kept[i]) bars[bin].kept++;
    });
    return bars;
}

/**
 * 分布を作る。値が 1 つも無いか、最小と最大が同じ（並べる幅が無い）ときは描かない
 */
export function distributionOf(rows: RowValue[], kept: boolean[], min: number | null, max: number | null, shape: Shape): Distribution {
    if (min === null || max === null || min === max) return { kind: "none" };
    const valued = rows.filter((r) => r.value !== null).length;
    if (valued === 0) return { kind: "none" };
    if (shape !== SHAPES.bars) {
        const points = pointsOf(rows, kept, min, max);
        const fits = valued <= POINT_LIMIT && points.every((p) => p.stack < POINT_STACK_LIMIT);
        if (shape === SHAPES.points || fits) return { kind: "points", points };
    }
    const bars = barsOf(rows, kept, min, max);
    return { kind: "bars", bars, maxCount: maxOf(bars.map((b) => b.count)) };
}
