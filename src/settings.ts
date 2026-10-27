"use strict";

/**
 * 書式ペイン。カードの name とスライスの name は capabilities.json の objects と完全に一致させる
 * （ずれてもエラーは出ず、書式ペインの変更が値に反映されないだけになる。format-pane Skill）。
 *
 * カードと既定値は、標準の数値の範囲スライサー（スタイル「指定の値の間」）の「スライサーの設定」「スライサー ヘッダー」「値」「スライダー」に合わせる。
 * 標準のヘッダーにある「保留中のアイコン」は、適用ボタンを待つ間の印で、確定したらすぐ掛けるこのスライサーには無いので置かない。
 *
 * 範囲の端・上位と下位の件数（objects.range）と絞り込み（objects.general.filter）は書式ペインに出さない。閲覧者がビジュアルの中で動かす。
 */

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import { NumberRange, applyNumberRanges, fontControl } from "./shared/formatCards";
import { dropdownValue } from "./shared/formatting";
import { FontValues, fontValuesOf } from "./shared/textStyle";
import { MODES, Mode, modeOf } from "./condition";
import { SHAPES, Shape, shapeOf } from "./distribution";

import CompositeCard = formattingSettings.CompositeCard;
import Group = formattingSettings.Group;
import Model = formattingSettings.Model;

/** 標準のスライサーの既定（2026-10 の Desktop、基本テーマで読んだ値） */
export const DEFAULTS = {
    headerFont: { family: "DIN", size: 12 },
    headerColor: "#252423",
    headerBorderColor: "#605E5C",
    headerBorderWidth: 1,
    valuesFont: { family: "Segoe UI", size: 9 },
    valuesColor: "#252423",
    handleFill: "#FFFFFF",
    handleBorder: "#605E5C",
    handleSize: 26,
    selectedColor: "#605E5C",
    trackColor: "#F3F2F1",
    /** 分布（標準に無い）。通る値はスライダーの選択した範囲と同じ灰、外れた値は薄い灰 */
    distributionKept: "#605E5C",
    distributionDropped: "#D2D0CE",
};

/** 数の欄の範囲。描くときも同じ範囲に丸める（lookOf） */
export const NUMBER_RANGES: Readonly<Record<string, NumberRange>> = {
    "header.borderWidth": [1, 10],
    "slider.handleSize": [10, 40],
};

export const STYLE_ITEMS: powerbi.IEnumMember[] = [
    { value: MODES.range, displayName: "範囲" },
    { value: MODES.top, displayName: "上位 N" },
    { value: MODES.bottom, displayName: "下位 N" },
];

const color = (name: string, displayName: string, value: string) => new formattingSettings.ColorPicker({ name, displayName, value: { value } });
const toggle = (name: string, displayName: string, value: boolean) => new formattingSettings.ToggleSwitch({ name, displayName, value });

export class SlicerSettingsCardSettings extends CompositeCard {
    name = "slicerSettings";
    displayName = "スライサーの設定";

    style = new formattingSettings.ItemDropdown({
        name: "style",
        displayName: "スタイル",
        description: "範囲はメジャーの値の下限〜上限で、上位 N・下位 N はメジャーの順位で絞る",
        items: STYLE_ITEMS,
        value: STYLE_ITEMS[0],
    });

    exclude = new formattingSettings.ToggleSwitch({
        name: "exclude",
        displayName: "合うものを除く",
        description: "オンにすると、条件に合う値を外し、残りでほかのビジュアルを絞る",
        value: false,
    });

    optionsGroup = new Group({ name: "slicerOptions", displayName: "オプション", slices: [this.style, this.exclude] });
    groups = [this.optionsGroup];

    get mode(): Mode {
        return modeOf(dropdownValue(this.style, MODES.range));
    }
}

/** スライサー ヘッダー。保存先は 1.1 までの「見出し」と同じ header（文字のサイズは px から pt に変わった） */
export class HeaderCardSettings extends CompositeCard {
    name = "header";
    displayName = "スライサー ヘッダー";

    show = toggle("show", "スライサー ヘッダー", true);
    topLevelSlice = this.show;

    titleText = new formattingSettings.TextInput({ name: "titleText", displayName: "タイトル テキスト", value: "", placeholder: "メジャーの名前" });
    font = fontControl({ controlName: "headerFont", displayName: "フォント", size: DEFAULTS.headerFont.size, family: DEFAULTS.headerFont.family });
    fontColor = color("fontColor", "フォントの色", DEFAULTS.headerColor);

    borderTop = toggle("borderTop", "上", false);
    borderBottom = toggle("borderBottom", "下", false);
    borderLeft = toggle("borderLeft", "左", false);
    borderRight = toggle("borderRight", "右", false);
    borderColor = color("borderColor", "カラー", DEFAULTS.headerBorderColor);
    borderWidth = new formattingSettings.NumUpDown({ name: "borderWidth", displayName: "線の幅", value: DEFAULTS.headerBorderWidth });

    background = color("background", "カラー", "");

    textGroup = new Group({ name: "headerText", displayName: "テキスト", slices: [this.titleText, this.font, this.fontColor] });
    borderGroup = new Group({
        name: "headerBorder",
        displayName: "罫線",
        slices: [this.borderTop, this.borderBottom, this.borderLeft, this.borderRight, this.borderColor, this.borderWidth],
    });
    backgroundGroup = new Group({ name: "headerBackground", displayName: "背景", slices: [this.background] });
    groups = [this.textGroup, this.borderGroup, this.backgroundGroup];
}

/** 値（下限・上限と件数の入力欄） */
export class ValuesCardSettings extends CompositeCard {
    name = "values";
    displayName = "値";

    font = fontControl({ controlName: "valuesFont", displayName: "フォント", size: DEFAULTS.valuesFont.size, family: DEFAULTS.valuesFont.family });
    fontColor = color("fontColor", "フォントの色", DEFAULTS.valuesColor);
    background = color("background", "背景色", "");

    valuesGroup = new Group({ name: "valuesText", displayName: "値", slices: [this.font, this.fontColor] });
    backgroundGroup = new Group({ name: "valuesBackground", displayName: "背景", slices: [this.background] });
    groups = [this.valuesGroup, this.backgroundGroup];
}

/** つまみの上の現在値（標準に無い）を出すとき */
export const VALUE_LABELS = { drag: "drag", always: "always", off: "off" } as const;
export type ValueLabel = (typeof VALUE_LABELS)[keyof typeof VALUE_LABELS];

export const VALUE_LABEL_ITEMS: powerbi.IEnumMember[] = [
    { value: VALUE_LABELS.drag, displayName: "動かしている間" },
    { value: VALUE_LABELS.always, displayName: "常に" },
    { value: VALUE_LABELS.off, displayName: "表示しない" },
];

export const SHAPE_ITEMS: powerbi.IEnumMember[] = [
    { value: SHAPES.auto, displayName: "自動" },
    { value: SHAPES.bars, displayName: "棒" },
    { value: SHAPES.points, displayName: "点" },
];

function valueLabelOf(value: unknown): ValueLabel {
    return value === VALUE_LABELS.always || value === VALUE_LABELS.off ? value : VALUE_LABELS.drag;
}

export class SliderCardSettings extends CompositeCard {
    name = "slider";
    displayName = "スライダー";

    show = toggle("show", "スライダー", true);
    topLevelSlice = this.show;

    handleFill = color("handleFill", "塗りつぶしの色", DEFAULTS.handleFill);
    handleBorder = color("handleBorder", "罫線", DEFAULTS.handleBorder);
    handleSize = new formattingSettings.NumUpDown({ name: "handleSize", displayName: "サイズ (px)", value: DEFAULTS.handleSize });
    selectedColor = color("selectedColor", "カラー", DEFAULTS.selectedColor);
    trackColor = color("trackColor", "カラー", DEFAULTS.trackColor);

    handleGroup = new Group({ name: "sliderHandle", displayName: "ハンドル", slices: [this.handleFill, this.handleBorder, this.handleSize] });
    selectedGroup = new Group({ name: "sliderSelected", displayName: "選択した範囲", slices: [this.selectedColor] });
    trackGroup = new Group({ name: "sliderTrack", displayName: "追跡", slices: [this.trackColor] });
    valueLabel = new formattingSettings.ItemDropdown({
        name: "valueLabel",
        displayName: "表示",
        description: "つまみの上に、つまみの位置の値を出す",
        items: VALUE_LABEL_ITEMS,
        value: VALUE_LABEL_ITEMS[0],
    });
    valueLabelGroup = new Group({ name: "sliderValueLabel", displayName: "現在値", slices: [this.valueLabel] });
    groups = [this.handleGroup, this.selectedGroup, this.trackGroup, this.valueLabelGroup];
}

/** 分布（標準に無い）。スライダーの上に、メジャーの値がどこに何件あるかを棒か点で描く */
export class DistributionCardSettings extends CompositeCard {
    name = "distribution";
    displayName = "分布";

    show = toggle("show", "分布", false);
    topLevelSlice = this.show;

    shape = new formattingSettings.ItemDropdown({
        name: "shape",
        displayName: "形",
        description: "自動は、件数が少なければ 1 件 1 点、多ければ棒",
        items: SHAPE_ITEMS,
        value: SHAPE_ITEMS[0],
    });
    keptColor = color("keptColor", "条件に合う値", DEFAULTS.distributionKept);
    droppedColor = color("droppedColor", "条件に合わない値", DEFAULTS.distributionDropped);

    optionsGroup = new Group({ name: "distributionOptions", displayName: "オプション", slices: [this.shape] });
    colorGroup = new Group({ name: "distributionColors", displayName: "色", slices: [this.keptColor, this.droppedColor] });
    groups = [this.optionsGroup, this.colorGroup];
}

export class VisualFormattingSettingsModel extends Model {
    slicerSettings = new SlicerSettingsCardSettings();
    header = new HeaderCardSettings();
    values = new ValuesCardSettings();
    slider = new SliderCardSettings();
    distribution = new DistributionCardSettings();
    cards = [this.slicerSettings, this.header, this.values, this.slider, this.distribution];

    constructor() {
        super();
        applyNumberRanges(this.cards, NUMBER_RANGES);
    }
}

/** 画面に渡す見た目。色は空なら塗らない（背景）か既定 */
export interface Look {
    /** ハイコントラスト。色が 2 色だけになるので、選んだ範囲を線の太さで分ける */
    highContrast: boolean;
    header: {
        show: boolean;
        /** 空ならメジャーの名前 */
        text: string;
        font: FontValues;
        color: string;
        border: { top: boolean; bottom: boolean; left: boolean; right: boolean; color: string; width: number };
        background: string;
    };
    /** border は入力欄の枠。空なら操作の部品の既定（ハイコントラストでは前景色） */
    values: { font: FontValues; color: string; background: string; border: string };
    slider: { show: boolean; handleFill: string; handleBorder: string; handleSize: number; selected: string; track: string; valueLabel: ValueLabel };
    /** 分布。ハイコントラストでは、合う値を前景色で塗り、合わない値は前景色の枠だけにする（outline） */
    distribution: { show: boolean; shape: Shape; kept: string; dropped: string; outline: boolean };
}

const colorOf = (picker: formattingSettings.ColorPicker, fallback: string) => picker.value?.value || fallback;
const clamp = (value: number | undefined, [min, max]: NumberRange, fallback: number) => Math.min(max, Math.max(min, value ?? fallback));

/**
 * 書式の設定から見た目を作る。ハイコントラストでは、選んだ色を使わず前景色と背景色の 2 色にする
 * （つまみの塗りと背景は背景色、ほかは前景色）
 */
export function lookOf(settings: VisualFormattingSettingsModel, highContrast?: { foreground: string; background: string }): Look {
    const { header, values, slider, distribution } = settings;
    const fg = highContrast?.foreground;
    const bg = highContrast?.background;
    return {
        highContrast: !!highContrast,
        header: {
            show: header.show.value,
            text: header.titleText.value?.trim() ?? "",
            font: fontValuesOf(header.font, DEFAULTS.headerFont.family, DEFAULTS.headerFont.size),
            color: fg ?? colorOf(header.fontColor, DEFAULTS.headerColor),
            border: {
                top: header.borderTop.value,
                bottom: header.borderBottom.value,
                left: header.borderLeft.value,
                right: header.borderRight.value,
                color: fg ?? colorOf(header.borderColor, DEFAULTS.headerBorderColor),
                width: clamp(header.borderWidth.value, NUMBER_RANGES["header.borderWidth"], DEFAULTS.headerBorderWidth),
            },
            background: bg ?? colorOf(header.background, ""),
        },
        values: {
            font: fontValuesOf(values.font, DEFAULTS.valuesFont.family, DEFAULTS.valuesFont.size),
            color: fg ?? colorOf(values.fontColor, DEFAULTS.valuesColor),
            background: bg ?? colorOf(values.background, ""),
            border: fg ?? "",
        },
        slider: {
            show: slider.show.value,
            handleFill: bg ?? colorOf(slider.handleFill, DEFAULTS.handleFill),
            handleBorder: fg ?? colorOf(slider.handleBorder, DEFAULTS.handleBorder),
            handleSize: clamp(slider.handleSize.value, NUMBER_RANGES["slider.handleSize"], DEFAULTS.handleSize),
            selected: fg ?? colorOf(slider.selectedColor, DEFAULTS.selectedColor),
            track: fg ?? colorOf(slider.trackColor, DEFAULTS.trackColor),
            valueLabel: valueLabelOf(dropdownValue(slider.valueLabel, VALUE_LABELS.drag)),
        },
        distribution: {
            show: distribution.show.value,
            shape: shapeOf(dropdownValue(distribution.shape, SHAPES.auto)),
            kept: fg ?? colorOf(distribution.keptColor, DEFAULTS.distributionKept),
            dropped: fg ?? colorOf(distribution.droppedColor, DEFAULTS.distributionDropped),
            outline: !!highContrast,
        },
    };
}

/** 書式を何も変えていないときの見た目（テストとハーネス） */
export const DEFAULT_LOOK: Look = lookOf(new VisualFormattingSettingsModel());
