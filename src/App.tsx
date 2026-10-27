"use strict";

/**
 * メジャーのスライサーの画面。見出し（メジャーの名前）、条件の入力、件数。
 * 条件の入力は、範囲なら下限・上限の数の入力とつまみ 2 つのスライダー、上位 N・下位 N なら件数の入力。
 * 形と既定の色は標準の数値の範囲スライサー（指定の値の間）に合わせ、見出しには標準の消しゴムと同じ解除のボタンを置く。
 *
 * 動かしている間は手元の状態だけを変え、放したとき（入力は Enter か外へ移ったとき）に onCommit で確定する。
 * 確定のたびに絞り込みを掛け直すので、つまみを動かすたびにページ全体を描き直さないため。
 */

import * as React from "react";
import { OPEN_RANGE, Range, parseNumberInput } from "./range";
import { Condition, MODES, isUnset, keptRows } from "./condition";
import { ViewModel } from "./viewModel";
import { Look, VALUE_LABELS } from "./settings";
import { Distribution, POINT_STACK_LIMIT, distributionOf } from "./distribution";
import { fontStyleOf } from "./shared/textStyle";
import { BLANK_TEXT } from "./shared/tooltip";

/** スライダーの刻みの数。つまみの位置は 0〜STEPS の整数で持ち、最小〜最大に割り当てる */
export const SLIDER_STEPS = 1000;

export const EMPTY_GUIDE = "「フィールド」と「値」を入れてください";
export const NO_MATCH_TEXT = "該当なし";
export const CLEAR_LABEL = "選択のクリア";

export interface AppProps {
    viewModel: ViewModel;
    condition: Condition;
    /** 操作の部品の CSS 変数（controlVars） */
    colorVars: Record<string, string>;
    look: Look;
    interactive: boolean;
    editing: boolean;
    /** 分布の点のツールヒント（Power BI の tooltipService）。無ければブラウザーのツールヒント（title）で出す */
    tooltip?: PointTooltip;
    /** 範囲か件数を確定した（変えたものだけを渡す。解除のボタンは両方を空で渡す） */
    onCommit: (next: Partial<Pick<Condition, "range" | "count">>) => void;
}

/** ツールヒントの 1 行（Power BI の VisualTooltipDataItem と同じ形） */
export interface TooltipItem {
    displayName: string;
    value: string;
}

export interface PointTooltip {
    show: (items: TooltipItem[], clientX: number, clientY: number, move: boolean) => void;
    hide: () => void;
}

/** 点のツールヒントの中身。標準のビジュアルと同じく、列の名前と値、メジャーの名前と値の 2 行 */
export function pointTooltipItems(viewModel: ViewModel, row: number): TooltipItem[] {
    const r = viewModel.rows[row];
    return [
        { displayName: viewModel.columnName, value: String(r.key ?? BLANK_TEXT) },
        { displayName: viewModel.measureName, value: r.value === null ? "" : viewModel.format(r.value) },
    ];
}

function pointTooltipProps(items: TooltipItem[], tooltip: PointTooltip | undefined): React.HTMLAttributes<HTMLSpanElement> {
    if (!tooltip) return { title: items.map((i) => `${i.displayName}：${i.value}`).join("\n") };
    return {
        onMouseEnter: (e) => tooltip.show(items, e.clientX, e.clientY, false),
        onMouseMove: (e) => tooltip.show(items, e.clientX, e.clientY, true),
        onMouseLeave: () => tooltip.hide(),
    };
}

/** 件数の文。条件を決めていなければ全件、通る値が無ければ「該当なし」。除外のときは、除いたあとに残る数 */
export function statusText(viewModel: ViewModel, condition: Condition): string {
    const total = viewModel.rows.length;
    if (isUnset(condition)) return `全 ${total.toLocaleString("en-US")} 件`;
    const matched = keptRows(viewModel.rows, condition).filter(Boolean).length;
    if (matched === 0) return NO_MATCH_TEXT;
    return `該当 ${matched.toLocaleString("en-US")} / ${total.toLocaleString("en-US")} 件`;
}

function positionOf(value: number | null, min: number, max: number, fallback: number): number {
    if (value === null || max === min) return fallback;
    const p = Math.round(((value - min) / (max - min)) * SLIDER_STEPS);
    return Math.min(SLIDER_STEPS, Math.max(0, p));
}

/** つまみの位置を値に戻す。端まで寄せたら開いた端（null） */
export function valueAtPosition(position: number, min: number, max: number, end: "from" | "to"): number | null {
    if (end === "from" && position <= 0) return null;
    if (end === "to" && position >= SLIDER_STEPS) return null;
    return min + ((max - min) * position) / SLIDER_STEPS;
}

interface BoundInputProps {
    value: number | null;
    /** 空のときに値として出す文字（標準と同じく、データの最小・最大）。そのまま確定しても空のまま */
    fallback: string;
    placeholder: string;
    style: React.CSSProperties;
    label: string;
    disabled: boolean;
    format: (value: number) => string;
    percent: boolean;
    onCommit: (value: number | null) => void;
    /** 打った文字の読み方。既定は parseNumberInput（null は空、undefined は読めない） */
    parse?: (text: string) => number | null | undefined;
}

/**
 * 下限・上限の入力。見えている書式（10.0% など）で出し、打った数は parseNumberInput で読む。読めなければ元に戻す。
 * 選んだら文字を全部選ぶ（標準と同じ。打てばそのまま置き換わる）
 */
function BoundInput({ value, fallback, placeholder, style, label, disabled, format, percent, onCommit, parse }: BoundInputProps): React.ReactElement {
    const shown = value === null ? fallback : format(value);
    const [text, setText] = React.useState(shown);
    const [focused, setFocused] = React.useState(false);
    React.useEffect(() => {
        if (!focused) setText(shown);
    }, [shown, focused]);

    // Esc で戻したあとの blur では確定しない（戻した文字が state に入る前に blur が走り、打った文字で確定していた）
    const cancelled = React.useRef(false);
    const commit = () => {
        if (cancelled.current) {
            cancelled.current = false;
            setText(shown);
            return;
        }
        const next = inputValue(text, value, fallback, parse ?? ((t) => parseNumberInput(t, percent)));
        if (next === undefined) {
            setText(shown);
            return;
        }
        if (next !== value) onCommit(next);
    };

    return (
        <input
            className="ms-input"
            style={style}
            type="text"
            inputMode="decimal"
            aria-label={label}
            value={text}
            placeholder={placeholder}
            disabled={disabled}
            onFocus={(e) => {
                setFocused(true);
                e.target.select();
            }}
            onChange={(e) => setText(e.target.value)}
            onBlur={() => {
                setFocused(false);
                commit();
            }}
            onKeyDown={(e) => {
                if (e.key === "Enter") (e.target as HTMLInputElement).blur();
                if (e.key === "Escape") {
                    cancelled.current = true;
                    setText(shown);
                    (e.target as HTMLInputElement).blur();
                }
            }}
        />
    );
}

/** 標準の消しゴム（選択のクリア）に似せた絵。16px の枠に描く */
function EraserIcon(): React.ReactElement {
    return (
        <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
            <path
                d="M9.2 2.3a1 1 0 0 1 1.4 0l3.1 3.1a1 1 0 0 1 0 1.4L8.4 12.1H13v1H5.6L2.3 9.8a1 1 0 0 1 0-1.4zM5.3 8.6 3 10.9 5.9 13.1h1.1l1.6-1.6z"
                fill="none"
                stroke="currentColor"
                strokeWidth="1"
                strokeLinejoin="round"
            />
        </svg>
    );
}

/** つまみが重なって見える近さ（刻み）。幅 300 のビジュアルで、つまみ 26px が線のおよそ 1 割 */
const THUMB_OVERLAP = SLIDER_STEPS * 0.1;

/** つまみの中心の位置（線はつまみの半径だけ内側にある）。ハンドルの大きさは CSS 変数から取る */
function thumbLeft(position: number): string {
    return `calc(var(--ms-handle-size) / 2 + (100% - var(--ms-handle-size)) * ${position / SLIDER_STEPS})`;
}

/** 2 つの吹き出しが重なる近さ。これより近ければ「下限 – 上限」の 1 つにまとめる */
const LABEL_MERGE = SLIDER_STEPS * 0.22;

/** 吹き出しを端に寄せる位置。線の端の 15% より外側では、吹き出しをつまみの中心から内側へずらす（ビジュアルの枠で切れないように） */
const LABEL_EDGE = SLIDER_STEPS * 0.15;

function labelClass(position: number): string {
    if (position < LABEL_EDGE) return "ms-value-label ms-start";
    if (position > SLIDER_STEPS - LABEL_EDGE) return "ms-value-label ms-end";
    return "ms-value-label";
}

export interface ValueLabelEnd {
    position: number;
    text: string;
    /** 出すか（動かしている間は動かしたつまみだけ） */
    shown: boolean;
}

/**
 * つまみの上の現在値（標準に無い）。形は Web の部品（MUI・Ant Design・Mantine のスライダー）の吹き出しにならい、
 * 濃い地に白い字で、つまみを指す小さな三角を付ける。2 つが近ければ「下限 – 上限」の 1 つにまとめる
 */
export function ValueLabels({ from, to }: { from: ValueLabelEnd; to: ValueLabelEnd }): React.ReactElement | null {
    if (to.position - from.position < LABEL_MERGE && (from.shown || to.shown)) {
        const label = from.text === to.text ? from.text : `${from.text} – ${to.text}`;
        const center = (from.position + to.position) / 2;
        return (
            <div className={labelClass(center)} style={{ left: thumbLeft(center) }}>
                {label}
            </div>
        );
    }
    return (
        <>
            {[from, to].map((end, i) =>
                end.shown ? (
                    <div key={i} className={labelClass(end.position)} style={{ left: thumbLeft(end.position) }}>
                        {end.text}
                    </div>
                ) : null
            )}
        </>
    );
}

/** 点の大きさと段の高さ（px） */
const POINT_SIZE = 7;
const POINT_STEP = 8;

/**
 * 分布の帯（標準に無い）。スライダーの線と同じ左右の位置に、棒か点で描く。
 * 点に乗せると列の値とメジャーの値が出る（Power BI のツールヒント）。乗せて読む部品で、操作はしない
 */
function DistributionStrip({ distribution, viewModel, outline, tooltip }: { distribution: Distribution; viewModel: ViewModel; outline: boolean; tooltip?: PointTooltip }): React.ReactElement | null {
    if (distribution.kind === "none") return null;
    if (distribution.kind === "points") {
        return (
            <div className="ms-dist" aria-hidden="true">
                {distribution.points.map((p) => {
                    return (
                        <span
                            key={p.row}
                            className={p.kept ? "ms-point ms-kept" : "ms-point"}
                            style={{ left: `${p.position * 100}%`, bottom: `${Math.min(p.stack, POINT_STACK_LIMIT - 1) * POINT_STEP}px`, width: POINT_SIZE, height: POINT_SIZE }}
                            {...pointTooltipProps(pointTooltipItems(viewModel, p.row), tooltip)}
                        />
                    );
                })}
            </div>
        );
    }
    const width = 100 / distribution.bars.length;
    return (
        <div className={outline ? "ms-dist ms-outline" : "ms-dist"} aria-hidden="true">
            {distribution.bars.map((b, i) =>
                b.count === 0 ? null : (
                    <span
                        key={i}
                        className="ms-bar"
                        style={{ left: `${b.start * 100}%`, width: `calc(${width}% - 2px)`, height: `${(b.count / distribution.maxCount) * 100}%` }}
                    >
                        {b.kept > 0 && <span className="ms-bar-kept" style={{ height: `${(b.kept / b.count) * 100}%` }} />}
                    </span>
                )
            )}
        </div>
    );
}

export function App({ viewModel, condition, colorVars, look, interactive, editing, tooltip, onCommit }: AppProps): React.ReactElement {
    const { min, max } = viewModel;
    const range = condition.range;
    const hasScale = min !== null && max !== null;
    const lo = min ?? 0;
    const hi = max ?? 0;
    const committedFrom = positionOf(range.from, lo, hi, 0);
    const committedTo = positionOf(range.to, lo, hi, SLIDER_STEPS);
    // 動かしている間のつまみの位置と、どちらのつまみを動かしたか。放したら確定して、確定した値から作り直す
    const [drag, setDrag] = React.useState<{ from: number; to: number; movedFrom: boolean; movedTo: boolean } | null>(null);
    const from = drag?.from ?? committedFrom;
    const to = drag?.to ?? committedTo;
    // つまみが重なると上にある上限のつまみしか掴めない。重なったまま下限の側へ動かしたら、その操作のあいだは下限を動かす
    // （逆も同じ。2026-10-27 の Desktop で、右端で重なると下限をマウスで動かせなかった）
    const steer = React.useRef<"from" | "to" | null>(null);
    const moveFrom = (position: number) =>
        setDrag({ from: Math.min(position, to), to, movedFrom: true, movedTo: drag?.movedTo ?? false });
    const moveTo = (position: number) =>
        setDrag({ from, to: Math.max(position, from), movedFrom: drag?.movedFrom ?? false, movedTo: true });

    // 件数の文が下で切れるときは、半分だけ見せずに隠す（低いビジュアル。標準には件数の文が無い）。
    // 大きさが変わると Power BI が update を呼び、描き直すたびにここで測り直す
    const rootRef = React.useRef<HTMLDivElement>(null);
    const statusRef = React.useRef<HTMLDivElement>(null);
    React.useLayoutEffect(() => {
        const root = rootRef.current;
        const line = statusRef.current;
        if (!root || !line) return;
        line.style.visibility = "";
        const fits = line.getBoundingClientRect().bottom <= root.getBoundingClientRect().bottom + 0.5;
        line.style.visibility = fits ? "" : "hidden";
    });

    const commitDrag = () => {
        steer.current = null;
        if (!drag) return;
        setDrag(null);
        onCommit({ range: draggedRange(drag, range, lo, hi) });
    };

    const rootStyle = {
        ...colorVars,
        "--ms-handle-size": `${look.slider.handleSize}px`,
        "--ms-handle-fill": look.slider.handleFill,
        "--ms-handle-border": look.slider.handleBorder,
        "--ms-selected": look.slider.selected,
        "--ms-track": look.slider.track,
        "--ms-dist-kept": look.distribution.kept,
        "--ms-dist-dropped": look.distribution.dropped,
        ...(look.values.border ? { "--ms-input-border": look.values.border } : {}),
    } as React.CSSProperties;

    if (!viewModel.ready) {
        return (
            <div className="ms-root" style={rootStyle}>
                {editing && <div className="ms-empty">{EMPTY_GUIDE}</div>}
            </div>
        );
    }

    const disabled = !interactive;
    const sliderDisabled = disabled || !hasScale || lo === hi;
    // 件数の文と分布は、動かしている間も今のつまみの位置で数える
    const shownCondition = drag ? { ...condition, range: draggedRange(drag, range, lo, hi) } : condition;
    // 吹き出しの値は件数の文と同じ範囲から出す（動かしていない端は打った値のまま。刻みに丸めない）。開いた端は最小・最大
    const shownRange = shownCondition.range;
    const labelText = (value: number | null, fallback: number) => viewModel.format(value ?? fallback);
    const status = statusText(viewModel, shownCondition);
    const distribution: Distribution = look.distribution.show
        ? distributionOf(viewModel.rows, keptRows(viewModel.rows, shownCondition), min, max, look.distribution.shape)
        : { kind: "none" };
    const showLabels = !sliderDisabled && (look.slider.valueLabel === VALUE_LABELS.always || (look.slider.valueLabel === VALUE_LABELS.drag && drag !== null));
    const ranked = condition.mode !== MODES.range;
    const headerText = look.header.text || viewModel.measureName;
    const border = look.header.border;
    const borderOf = (on: boolean) => (on ? `${border.width}px solid ${border.color}` : undefined);
    const valueStyle: React.CSSProperties = {
        ...fontStyleOf(look.values.font),
        color: look.values.color,
        background: look.values.background || undefined,
    };

    return (
        <div className={look.highContrast ? "ms-root ms-hc" : "ms-root"} style={rootStyle} ref={rootRef}>
            {look.header.show && (
                <div
                    className="ms-header"
                    style={{
                        background: look.header.background || undefined,
                        borderTop: borderOf(border.top),
                        borderBottom: borderOf(border.bottom),
                        borderLeft: borderOf(border.left),
                        borderRight: borderOf(border.right),
                    }}
                >
                    <span className="ms-header-text" style={{ ...fontStyleOf(look.header.font), color: look.header.color }} title={headerText}>
                        {headerText}
                    </span>
                    {interactive && !isUnset(condition) && (
                        <button
                            type="button"
                            className="ms-clear"
                            aria-label={CLEAR_LABEL}
                            title={CLEAR_LABEL}
                            onClick={() => onCommit({ range: OPEN_RANGE, count: null })}
                        >
                            <EraserIcon />
                        </button>
                    )}
                </div>
            )}
            {ranked ? (
                <>
                <div className="ms-inputs">
                    <span className="ms-label" style={valueStyle}>{condition.mode === MODES.top ? "上位" : "下位"}</span>
                    <BoundInput
                        value={condition.count}
                        fallback=""
                        placeholder="件数"
                        style={valueStyle}
                        label={`${viewModel.measureName}の${condition.mode === MODES.top ? "上位" : "下位"}の件数`}
                        disabled={disabled}
                        format={(v) => String(v)}
                        percent={false}
                        parse={parseCount}
                        onCommit={(count) => onCommit({ count })}
                    />
                    <span className="ms-label" style={valueStyle}>件</span>
                </div>
                <DistributionStrip distribution={distribution} viewModel={viewModel} outline={look.distribution.outline} tooltip={tooltip} />
                </>
            ) : (
                <>
                <div className="ms-inputs">
                    <BoundInput
                        value={range.from}
                        fallback={hasScale ? viewModel.format(lo) : ""}
                        placeholder=""
                        style={valueStyle}
                        label={`${viewModel.measureName}の下限`}
                        disabled={disabled}
                        format={viewModel.format}
                        percent={viewModel.percent}
                        onCommit={(value) => onCommit({ range: orderRange({ from: value, to: range.to }) })}
                    />
                    <BoundInput
                        value={range.to}
                        fallback={hasScale ? viewModel.format(hi) : ""}
                        placeholder=""
                        style={valueStyle}
                        label={`${viewModel.measureName}の上限`}
                        disabled={disabled}
                        format={viewModel.format}
                        percent={viewModel.percent}
                        onCommit={(value) => onCommit({ range: orderRange({ from: range.from, to: value }) })}
                    />
                </div>
                <DistributionStrip distribution={distribution} viewModel={viewModel} outline={look.distribution.outline} tooltip={tooltip} />
                {look.slider.show && (
                <div className="ms-slider">
                    {/* 線はつまみの半径だけ内側に置く（つまみの中心が動く範囲と同じ） */}
                    <div className="ms-rail">
                        <div className="ms-track" />
                        <div className="ms-fill" style={{ left: `${(from / SLIDER_STEPS) * 100}%`, right: `${100 - (to / SLIDER_STEPS) * 100}%` }} />
                    </div>
                    <input
                        type="range"
                        aria-label={`${viewModel.measureName}の下限`}
                        title="開始"
                        min={0}
                        max={SLIDER_STEPS}
                        value={from}
                        disabled={sliderDisabled}
                        onChange={(e) => {
                            const position = Number(e.target.value);
                            if (steer.current === "to" || (steer.current === null && to - from <= THUMB_OVERLAP && position > to)) {
                                steer.current = "to";
                                moveTo(position);
                            } else {
                                moveFrom(position);
                            }
                        }}
                        onPointerUp={commitDrag}
                        onKeyUp={commitDrag}
                        onBlur={commitDrag}
                    />
                    <input
                        type="range"
                        aria-label={`${viewModel.measureName}の上限`}
                        title="終了"
                        min={0}
                        max={SLIDER_STEPS}
                        value={to}
                        disabled={sliderDisabled}
                        onChange={(e) => {
                            const position = Number(e.target.value);
                            if (steer.current === "from" || (steer.current === null && to - from <= THUMB_OVERLAP && position < from)) {
                                steer.current = "from";
                                moveFrom(position);
                            } else {
                                moveTo(position);
                            }
                        }}
                        onPointerUp={commitDrag}
                        onKeyUp={commitDrag}
                        onBlur={commitDrag}
                    />
                    {showLabels && (
                        <ValueLabels
                            from={{ position: from, text: labelText(shownRange.from, lo), shown: drag === null || drag.movedFrom }}
                            to={{ position: to, text: labelText(shownRange.to, hi), shown: drag === null || drag.movedTo }}
                        />
                    )}
                </div>
                )}
                </>
            )}
            <div className="ms-status" ref={statusRef}>
                {status}
            </div>
        </div>
    );
}

/**
 * つまみを放したときの範囲。動かしたつまみの端だけをつまみの位置から作り、動かしていない端は確定した値のまま残す
 * （両端を刻みで作り直すと、打った下限 3,000 が上限を動かしただけで 2,995 に変わった）
 */
export function draggedRange(drag: { from: number; to: number; movedFrom: boolean; movedTo: boolean }, range: Range, min: number, max: number): Range {
    return {
        from: drag.movedFrom ? valueAtPosition(drag.from, min, max, "from") : range.from,
        to: drag.movedTo ? valueAtPosition(drag.to, min, max, "to") : range.to,
    };
}

/** 下限が上限を超えたら入れ替える */
function orderRange(range: Range): Range {
    if (range.from !== null && range.to !== null && range.from > range.to) return { from: range.to, to: range.from };
    return range;
}

/**
 * 入力欄を離れたときの値。空のときに出していた最小・最大（fallback）をそのまま離れたら空のまま（null）で、
 * 読めない文字は undefined（元に戻す）
 */
export function inputValue(
    text: string,
    value: number | null,
    fallback: string,
    parse: (text: string) => number | null | undefined
): number | null | undefined {
    if (value === null && fallback !== "" && text.trim() === fallback) return null;
    return parse(text);
}

/** 件数の入力。1 以上の整数だけ。空は未指定（null）、それ以外は読めない（undefined） */
export function parseCount(text: string): number | null | undefined {
    const parsed = parseNumberInput(text, false);
    if (parsed === null) return null;
    if (parsed === undefined || parsed < 1 || !Number.isInteger(parsed)) return undefined;
    return parsed;
}
