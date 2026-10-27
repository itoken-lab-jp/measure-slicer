"use strict";

/**
 * DataView から、列の値ごとのメジャーの値を取り出す。
 *
 * 列の値はほかの作品と同じく 30,000 件まで読む（top）。上限に達したら警告を出す。
 * そのときは読めた値だけで絞るので、読めなかった値は Not In なら残り、In なら落ちる。
 */

import powerbi from "powerbi-visuals-api";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import { ROW_LIMIT } from "./shared/warnings";
import { FilterTarget, RowValue, targetOf } from "./filter";
import { isPercentFormat } from "./range";

export const TRUNCATED_TITLE = "すべての値を読み込めていません";
export const TRUNCATED_NOTICE = `列の値が ${ROW_LIMIT.toLocaleString("en-US")} 件に達したので、先頭の ${ROW_LIMIT.toLocaleString("en-US")} 件で絞っています。フィルターで減らしてください。`;

export interface ViewModel {
    /** 列とメジャーがそろっているか */
    ready: boolean;
    rows: RowValue[];
    target: FilterTarget | null;
    columnName: string;
    measureName: string;
    /** メジャーの最小・最大（空の値は除く）。値が無ければ null */
    min: number | null;
    max: number | null;
    formatString: string;
    percent: boolean;
    format: (value: number) => string;
    truncated: boolean;
}

const EMPTY: ViewModel = {
    ready: false,
    rows: [],
    target: null,
    columnName: "",
    measureName: "",
    min: null,
    max: null,
    formatString: "",
    percent: false,
    format: (v) => String(v),
    truncated: false,
};

export function transform(dataView: powerbi.DataView | undefined): ViewModel {
    const categorical = dataView?.categorical;
    const category = categorical?.categories?.[0];
    const measure = categorical?.values?.find((c) => c.source?.roles?.measure) ?? categorical?.values?.[0];
    if (!category || !measure) {
        return {
            ...EMPTY,
            columnName: category?.source?.displayName ?? "",
            measureName: measure?.source?.displayName ?? "",
        };
    }

    const rows: RowValue[] = category.values.map((key, i) => {
        const raw = measure.values[i];
        const value = typeof raw === "number" && Number.isFinite(raw) ? raw : null;
        return { key, value };
    });
    let min: number | null = null;
    let max: number | null = null;
    for (const { value } of rows) {
        if (value === null) continue;
        if (min === null || value < min) min = value;
        if (max === null || value > max) max = value;
    }
    const formatString = valueFormatter.getFormatStringByColumn(measure.source) ?? "";
    // 表示単位（千・百万）は付けない。入力に打つ数と同じ桁で見せる
    const formatter = valueFormatter.create({ format: formatString || undefined });

    return {
        ready: true,
        rows,
        target: targetOf(category.source),
        columnName: category.source.displayName ?? "",
        measureName: measure.source.displayName ?? "",
        min,
        max,
        formatString,
        percent: isPercentFormat(formatString),
        format: (value) => formatter.format(value),
        truncated: rows.length >= ROW_LIMIT,
    };
}
