/*
 *  Power BI Visual
 *  Licensed under the MIT License.
 */
"use strict";

import powerbi from "powerbi-visuals-api";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { displayWarnings } from "./shared/warnings";
import { GuardedRenderer } from "./shared/renderGuard";
import { controlColorsOf, controlVars } from "./shared/controls";
import { showTooltipAt } from "./shared/tooltip";

import "./../style/visual.less";
import { App, PointTooltip } from "./App";
import { VisualFormattingSettingsModel, lookOf } from "./settings";
import { TRUNCATED_NOTICE, TRUNCATED_TITLE, transform, ViewModel } from "./viewModel";
import { OPEN_RANGE, rangeOf } from "./range";
import { filterOf, planFilter, sameFilter } from "./filter";
import { Condition, countOf, keptRows } from "./condition";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import FilterAction = powerbi.FilterAction;

export class Visual implements IVisual {
    private root: Root;
    /** 描画の守り。update では描き終えてからレンダリング イベントを送り、失敗したら前の絵を消して知らせを出す */
    private renderer: GuardedRenderer;
    private host: IVisualHost;
    private events: IVisualEventService;
    private tooltip: PointTooltip;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    /**
     * 前の update で見たもの。絞り込みが掛かっていたか（Power BI から届いた jsonFilters が空でなかったか）と、条件。
     * 条件はそのままなのに絞り込みだけが消えたら、外から外された（すべてのスライサーをクリア）とみなす
     */
    private previous: { filtered: boolean; condition: string } | undefined;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.events = options.host.eventService;
        const tooltipService = options.host.tooltipService;
        this.tooltip = {
            show: (items, clientX, clientY, move) => showTooltipAt(tooltipService, options.element, clientX, clientY, move, items, []),
            hide: () => tooltipService.hide({ isTouchEvent: false, immediately: false }),
        };
        this.formattingSettingsService = new FormattingSettingsService();
        this.root = createRoot(options.element);
        this.renderer = new GuardedRenderer(this.root);
    }

    public update(options: VisualUpdateOptions): void {
        // レンダリングイベントは認定要件。started と、描き終えてから finished（失敗なら failed）を対で 1 回ずつ送る
        this.renderer.update(this.events, options, () => {
            const dataView = options.dataViews?.[0];
            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);
            const viewModel = transform(dataView);
            const objects = dataView?.metadata?.objects as { range?: { from?: unknown; to?: unknown; count?: unknown } } | undefined;
            let condition: Condition = {
                mode: this.formattingSettings.slicerSettings.mode,
                range: rangeOf(objects),
                count: countOf(objects?.range?.count),
                exclude: this.formattingSettings.slicerSettings.exclude.value,
            };
            displayWarnings(this.host, [viewModel.truncated && { title: TRUNCATED_TITLE, detail: TRUNCATED_NOTICE }]);
            // 操作できない場所（ダッシュボードのタイルなど）では、範囲を動かさせない
            const allowInteractions = this.host.hostCapabilities?.allowInteractions !== false;

            // ほかのスライサーでデータが変わると、範囲に合う値も入れ替わる。今掛かっているものと違えば掛け直す。
            // ただし、掛かっていた絞り込みが条件を変えないまま消えたら、外から外された（すべてのスライサーをクリア）ので、
            // 掛け直さずに範囲を空にする（掛け直すと、Desktop ではクリアしても範囲が残った）
            const current = options.jsonFilters ?? [];
            const key = JSON.stringify([condition.mode, condition.range, condition.count, condition.exclude]);
            const clearedOutside = allowInteractions && this.previous?.filtered === true && this.previous.condition === key && current.length === 0;
            if (clearedOutside) {
                this.persistCondition({ ...condition, range: OPEN_RANGE, count: null });
                condition = { ...condition, range: OPEN_RANGE, count: null };
            } else if (allowInteractions) {
                this.syncFilter(viewModel, condition, current);
            }
            this.previous = { filtered: current.length > 0, condition: key };

            const palette = this.host.colorPalette as powerbi.extensibility.ISandboxExtendedColorPalette;
            const foreground = palette?.foreground?.value ?? "#252423";
            const background = palette?.background?.value ?? "#FFFFFF";
            const highContrast = !!palette?.isHighContrast;
            const colors = controlColorsOf(foreground, background, highContrast);

            this.renderer.paint(
                React.createElement(App, {
                    viewModel,
                    condition,
                    colorVars: controlVars(colors),
                    look: lookOf(this.formattingSettings, highContrast ? { foreground, background } : undefined),
                    interactive: allowInteractions,
                    // 閲覧中（viewMode が View＝0）は、空のときの案内を出さない
                    editing: options.viewMode === undefined || options.viewMode !== 0,
                    tooltip: this.tooltip,
                    onCommit: (next) => this.commit(viewModel, { ...condition, ...next }),
                })
            );
        });
    }

    /**
     * 閲覧者が範囲か件数を確定した。決めた値は merge で残し、空にした値は remove で消して、絞り込みを掛け直す。
     * replace で決めた値だけを渡すと、Desktop では渡さなかった値が消えずに残った（上限 99,999 を空にしても戻った）
     */
    private commit(viewModel: ViewModel, condition: Condition): void {
        this.persistCondition(condition);
        this.syncFilter(viewModel, condition, undefined);
    }

    /** 範囲と件数を残す。決めた値は merge、空にした値は remove */
    private persistCondition(condition: Condition): void {
        const values: Array<[string, number | null]> = [
            ["from", condition.range.from],
            ["to", condition.range.to],
            ["count", condition.count],
        ];
        const set: powerbi.DataViewObject = {};
        const cleared: powerbi.DataViewObject = {};
        for (const [name, value] of values) {
            if (value === null) cleared[name] = null;
            else set[name] = value;
        }
        const changes: powerbi.VisualObjectInstancesToPersist = {};
        if (Object.keys(set).length > 0) changes.merge = [{ objectName: "range", selector: null, properties: set }];
        if (Object.keys(cleared).length > 0) changes.remove = [{ objectName: "range", selector: null, properties: cleared }];
        this.host.persistProperties(changes);
    }

    /** 条件に合う値を列への絞り込みにして掛ける。今のものと同じなら何もしない（current が無いときは必ず掛ける） */
    private syncFilter(viewModel: ViewModel, condition: Condition, current: readonly unknown[] | undefined): void {
        // 列とメジャーがそろっていない、列の先が分からないときは触らない（入れ替え中に掛けてある絞り込みを消さない）
        if (!viewModel.ready || !viewModel.target) return;
        const filter = filterOf(viewModel.target, planFilter(viewModel.rows, keptRows(viewModel.rows, condition)));
        if (current !== undefined && sameFilter(current, filter)) return;
        if (filter === null) {
            this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
        } else {
            this.host.applyJsonFilter(filter as unknown as powerbi.IFilter, "general", "filter", FilterAction.merge);
        }
    }

    /** 書式設定ペインを開くたび / 値変更のたびに呼ばれる */
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    public destroy(): void {
        this.root?.unmount();
    }
}
