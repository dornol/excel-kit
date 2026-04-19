# Styling

> [Back to Index](index.md)

## Column Styling

```java
writer
    .column("Amount", p -> p.amount(), cfg -> cfg
        .type(ExcelDataType.DOUBLE)
        .format("#,##0.00")
        .alignment(HorizontalAlignment.RIGHT)
        .backgroundColor(ExcelColor.LIGHT_YELLOW)
        .bold(true)
        .fontSize(12))
    .write(data);
```

## Default Column Style

Set writer-level defaults inherited by all columns unless overridden:

```java
ExcelWriter.<Product>create()
    .defaultStyle(d -> d
        .fontName("Arial").fontSize(10)
        .alignment(HorizontalAlignment.LEFT).bold(true))
    .column("Name", p -> p.name())                        // inherits all defaults
    .column("Price", p -> p.price(), c -> c
        .bold(false).alignment(HorizontalAlignment.RIGHT)) // overrides
    .write(data);
```

## Font

```java
.column("Text", p -> p.text(), cfg -> cfg
    .fontName("Arial")          // font family
    .fontSize(12)               // size in points
    .bold(true)                 // bold
    .fontColor(ExcelColor.RED)  // preset color
    .fontColor(255, 0, 0)       // RGB
    .strikethrough()            // strike-through
    .underline())               // underline
```

## Background Color

```java
.column("Cell", p -> p.val(), cfg -> cfg
    .backgroundColor(ExcelColor.LIGHT_YELLOW))
```

Available presets: `WHITE`, `BLACK`, `LIGHT_GRAY`, `GRAY`, `DARK_GRAY`, `RED`, `GREEN`, `BLUE`, `YELLOW`, `ORANGE`, `LIGHT_RED`, `LIGHT_GREEN`, `LIGHT_BLUE`, `LIGHT_YELLOW`, `LIGHT_ORANGE`, `LIGHT_PURPLE`, `CORAL`, `STEEL_BLUE`, `FOREST_GREEN`, `GOLD`

## Alignment

```java
.column("Left", p -> p.val(), cfg -> cfg.alignment(HorizontalAlignment.LEFT))
.column("Right", p -> p.val(), cfg -> cfg.alignment(HorizontalAlignment.RIGHT))
.column("Center", p -> p.val(), cfg -> cfg.alignment(HorizontalAlignment.CENTER))
```

## Vertical Alignment

```java
.column("Top", p -> p.val(), cfg -> cfg.verticalAlignment(VerticalAlignment.TOP))
.column("Bottom", p -> p.val(), cfg -> cfg.verticalAlignment(VerticalAlignment.BOTTOM))
.column("Justify", p -> p.val(), cfg -> cfg.verticalAlignment(VerticalAlignment.JUSTIFY))
```

## Text Wrapping

Enabled by default. Disable to clip content at column width:

```java
.column("Code", p -> p.code(), cfg -> cfg.wrapText(false))
```

## Text Rotation

Rotate text from -90 to 90 degrees:

```java
.column("Rotated", p -> p.label(), cfg -> cfg.rotation(45))   // 45 degrees counter-clockwise
.column("Vertical", p -> p.code(), cfg -> cfg.rotation(90))   // straight up
.column("Clock", p -> p.note(), cfg -> cfg.rotation(-30))     // 30 degrees clockwise
```

## Cell Indentation

Indent level 0-250:

```java
.column("Sub-item", p -> p.item(), cfg -> cfg
    .indentation(2).alignment(HorizontalAlignment.LEFT))
```

## Width Control

```java
.column("Fixed", p -> p.val(), cfg -> cfg.width(20))       // fixed width (disables auto-resize)
.column("Min", p -> p.val(), cfg -> cfg.minWidth(10))       // minimum auto-resize width
.column("Max", p -> p.val(), cfg -> cfg.maxWidth(50))       // maximum auto-resize width
```

## Cell Border

```java
// Uniform border
.column("Amount", p -> p.amount(), cfg -> cfg.border(ExcelBorderStyle.MEDIUM))

// Per-side border
.column("Mixed", p -> p.value(), cfg -> cfg
    .borderTop(ExcelBorderStyle.THICK)
    .borderBottom(ExcelBorderStyle.THIN)
    .borderLeft(ExcelBorderStyle.DASHED)
    .borderRight(ExcelBorderStyle.DOTTED))

// Partial override: top=THICK, rest=MEDIUM
.column("Partial", p -> p.value(), cfg -> cfg
    .border(ExcelBorderStyle.MEDIUM)
    .borderTop(ExcelBorderStyle.THICK))
```

Styles: `NONE`, `THIN`, `MEDIUM`, `THICK`, `DASHED`, `DOTTED`, `DOUBLE`, `HAIR`, `MEDIUM_DASHED`, `DASH_DOT`

## Row-Level Styling

Conditional background colors for entire rows:

```java
writer
    .rowColor(p -> p.isError() ? ExcelColor.LIGHT_RED : null)
    .column("Name", p -> p.name())
    .write(data);
```

## Cell-Level Conditional Styling

Per-cell background via `CellColorFunction`:

```java
.column("Amount", p -> p.amount(), cfg -> cfg
    .type(ExcelDataType.DOUBLE)
    .cellColor((value, row) -> {
        double amt = ((Number) value).doubleValue();
        if (amt < 0) return ExcelColor.LIGHT_RED;
        if (amt > 10000) return ExcelColor.LIGHT_GREEN;
        return null;
    }))
```

**Priority:** `cellColor` > `rowColor` > column `backgroundColor`
