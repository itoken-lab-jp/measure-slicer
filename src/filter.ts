"use strict";

/**
 * 範囲に合う列の値を、その列への絞り込み（JSON のフィルター）に置き換える。
 *
 * カスタム ビジュアルが掛けられる絞り込みは列にだけ効く（メジャーには掛けられない）。だから
 * 「条件に合う列の値の一覧」を In で掛ける。通る値と落ちる値の少ないほうで出す（Not In）。
 * 合う値が 0 件のときは、何も通さない絞り込み（空白でも空白でなくもない）を掛け、ほかのビジュアルを空にする
 * （外すと「全部合っている」ように見える。2026-10-26 ユーザー決定）。
 */

import powerbi from "powerbi-visuals-api";

import PrimitiveValue = powerbi.PrimitiveValue;

/** 絞り込みを掛ける列 */
export interface FilterTarget {
    table: string;
    column: string;
}

/** 列の値 1 つ。key は絞り込みに渡す元の値 */
export interface RowValue {
    key: PrimitiveValue;
    value: number | null;
}

export type FilterPlan =
    | { kind: "none" }
    | { kind: "nothing" }
    | { kind: "in" | "notIn"; values: PrimitiveValue[] };

/**
 * 行ごとの通る・通らない（keptRows）から、掛ける絞り込みの形を決める。全部通るなら掛けない（条件を決めていないときも全部通る）。
 * 通る値と落ちる値の少ないほうで出す（同じ数なら In）
 */
export function planFilter(rows: RowValue[], kept: boolean[]): FilterPlan {
    if (rows.length === 0) return { kind: "none" };
    const passed: PrimitiveValue[] = [];
    const dropped: PrimitiveValue[] = [];
    rows.forEach((row, i) => (kept[i] ? passed : dropped).push(row.key));
    if (dropped.length === 0) return { kind: "none" };
    if (passed.length === 0) return { kind: "nothing" };
    return passed.length <= dropped.length ? { kind: "in", values: passed } : { kind: "notIn", values: dropped };
}

const BASIC_SCHEMA = "https://powerbi.com/product/schema#basic";
const ADVANCED_SCHEMA = "https://powerbi.com/product/schema#advanced";
/** powerbi-models の FilterType（Advanced = 0、Basic = 1） */
const FILTER_TYPE_ADVANCED = 0;
const FILTER_TYPE_BASIC = 1;

export interface JsonFilter {
    $schema: string;
    target: FilterTarget;
    filterType: number;
    operator?: "In" | "NotIn";
    values?: PrimitiveValue[];
    requireSingleSelection?: boolean;
    logicalOperator?: "And";
    conditions?: Array<{ operator: "IsBlank" | "IsNotBlank" }>;
}

export function filterOf(target: FilterTarget, plan: FilterPlan): JsonFilter | null {
    switch (plan.kind) {
        case "none":
            return null;
        case "nothing":
            return {
                $schema: ADVANCED_SCHEMA,
                target,
                filterType: FILTER_TYPE_ADVANCED,
                logicalOperator: "And",
                conditions: [{ operator: "IsBlank" }, { operator: "IsNotBlank" }],
            };
        default:
            return {
                $schema: BASIC_SCHEMA,
                target,
                filterType: FILTER_TYPE_BASIC,
                operator: plan.kind === "in" ? "In" : "NotIn",
                values: plan.values,
                requireSingleSelection: false,
            };
    }
}

/** 比べるための形。値の並びと、ホストが足す項目の違いは見ない */
function signatureOf(filter: unknown): string {
    if (!filter || typeof filter !== "object") return "";
    const f = filter as Partial<JsonFilter>;
    const values = Array.isArray(f.values) ? f.values.map((v) => JSON.stringify(v instanceof Date ? v.toISOString() : v)).sort() : [];
    const conditions = Array.isArray(f.conditions) ? f.conditions.map((c) => c?.operator).join(",") : "";
    return JSON.stringify([f.target?.table, f.target?.column, f.filterType, f.operator ?? "", values, f.logicalOperator ?? "", conditions]);
}

/** 今掛かっている絞り込み（update の jsonFilters）が、掛けたいものと同じか。同じなら掛け直さない（update が回り続けない） */
export function sameFilter(current: readonly unknown[] | undefined, next: JsonFilter | null): boolean {
    const live = (current ?? []).filter((f) => f && typeof f === "object");
    if (next === null) return live.length === 0;
    return live.length === 1 && signatureOf(live[0]) === signatureOf(next);
}

/**
 * 列の名前から絞り込みの先を作る。expr（テーブルと列の参照）があればそれを使う（テーブル名に「.」があっても割れない）。
 * 無ければ queryName（「テーブル.列」）を最初の「.」で割る
 */
export function targetOf(source: powerbi.DataViewMetadataColumn | undefined): FilterTarget | null {
    if (!source) return null;
    const expr = source.expr as { source?: { entity?: unknown }; ref?: unknown } | undefined;
    if (typeof expr?.source?.entity === "string" && typeof expr.ref === "string") {
        return { table: expr.source.entity, column: expr.ref };
    }
    const name = source.queryName ?? "";
    const dot = name.indexOf(".");
    if (dot <= 0 || dot === name.length - 1) return null;
    return { table: name.slice(0, dot), column: name.slice(dot + 1) };
}
