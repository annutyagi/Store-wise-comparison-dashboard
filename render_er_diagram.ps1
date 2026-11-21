Add-Type -AssemblyName System.Drawing

$width = 1920
$height = 1080
$bmp = New-Object System.Drawing.Bitmap($width, $height)
$g = [System.Drawing.Graphics]::FromImage($bmp)

$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

# Background
$bgBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#F8FAFC"))
$g.FillRectangle($bgBrush, 0, 0, $width, $height)

# Colors
$primaryColor = [System.Drawing.ColorTranslator]::FromHtml("#1A365D")
$secondaryColor = [System.Drawing.ColorTranslator]::FromHtml("#2B6CB0")
$charcoal = [System.Drawing.ColorTranslator]::FromHtml("#2D3748")
$muted = [System.Drawing.ColorTranslator]::FromHtml("#718096")
$cardBg = [System.Drawing.ColorTranslator]::FromHtml("#FFFFFF")
$cardBorder = [System.Drawing.ColorTranslator]::FromHtml("#CBD5E0")
$bannerBg = [System.Drawing.ColorTranslator]::FromHtml("#EDF2F7")
$lineColor = [System.Drawing.ColorTranslator]::FromHtml("#4A5568")

$penBorder = New-Object System.Drawing.Pen($cardBorder, [float]1.5)
$penLine = New-Object System.Drawing.Pen($lineColor, [float]2.5)

# Fonts with explicit [float] sizes
$fontTitle = New-Object System.Drawing.Font("Segoe UI", [float]24, [System.Drawing.FontStyle]::Bold)
$fontSub = New-Object System.Drawing.Font("Segoe UI", [float]12, [System.Drawing.FontStyle]::Regular)
$fontTblHeader = New-Object System.Drawing.Font("Segoe UI", [float]13, [System.Drawing.FontStyle]::Bold)
$fontColName = New-Object System.Drawing.Font("Consolas", [float]10.5, [System.Drawing.FontStyle]::Bold)
$fontColType = New-Object System.Drawing.Font("Segoe UI", [float]9.5, [System.Drawing.FontStyle]::Regular)
$fontBadge = New-Object System.Drawing.Font("Segoe UI", [float]8.5, [System.Drawing.FontStyle]::Bold)
$fontVerb = New-Object System.Drawing.Font("Segoe UI", [float]10.5, [System.Drawing.FontStyle]::Bold)

# 1. Header Banner
$brushBanner = New-Object System.Drawing.SolidBrush($bannerBg)
$g.FillRectangle($brushBanner, 50, 40, 1820, 80)
$g.DrawRectangle($penBorder, 50, 40, 1820, 80)

$brushPrimary = New-Object System.Drawing.SolidBrush($primaryColor)
$brushMuted = New-Object System.Drawing.SolidBrush($muted)
$brushCharcoal = New-Object System.Drawing.SolidBrush($charcoal)
$brushWhite = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)

$g.DrawString("STORE WISE SALES COMPARISON DASHBOARD — ENTITY RELATIONSHIP (ER) DIAGRAM", $fontTitle, $brushPrimary, [float]70, [float]48)
$g.DrawString("Student: Anu Tyagi | Roll: 2500320140025 | MCA 2nd Year (Data Analytics), ABES | Mentor: Prof. Kalpna Dwivedi | Relational 3NF Architecture", $fontSub, $brushMuted, [float]72, [float]92)

# Helper function to draw table entity
function Draw-TableEntity {
    param(
        [int]$x, [int]$y, [int]$w, [string]$title, [string]$tableType, [array]$columns
    )
    $headerHeight = 44
    $rowHeight = 26
    $totalHeight = $headerHeight + ($columns.Count * $rowHeight) + 12
    
    # Shadow
    $shadowBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(25, 0, 0, 0))
    $g.FillRectangle($shadowBrush, $x + 4, $y + 4, $w, $totalHeight)
    
    # Body background & border
    $brushCard = New-Object System.Drawing.SolidBrush($cardBg)
    $g.FillRectangle($brushCard, $x, $y, $w, $totalHeight)
    $g.DrawRectangle($penBorder, $x, $y, $w, $totalHeight)
    
    # Header background
    $brushHdr = New-Object System.Drawing.SolidBrush($primaryColor)
    $g.FillRectangle($brushHdr, $x, $y, $w, $headerHeight)
    
    # Header text
    $g.DrawString($title, $fontTblHeader, $brushWhite, [float]($x + 14), [float]($y + 8))
    $fontTypeBadge = New-Object System.Drawing.Font("Segoe UI", [float]9, [System.Drawing.FontStyle]::Italic)
    $brushSubHdr = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#E2E8F0"))
    $g.DrawString("[$tableType]", $fontTypeBadge, $brushSubHdr, [float]($x + 14), [float]($y + 26))
    
    # Rows
    $curY = $y + $headerHeight + 6
    for ($i = 0; $i -lt $columns.Count; $i++) {
        $col = $columns[$i]
        $colName = $col[0]
        $colType = $col[1]
        $colKey = $col[2]
        
        # Zebra line
        if ($i % 2 -eq 1) {
            $zebraBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#F7FAFC"))
            $g.FillRectangle($zebraBrush, $x + 1, $curY - 2, $w - 2, $rowHeight)
        }
        
        # Key Badge
        if ($colKey -eq "PK") {
            $badgeBg = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#FED7D7"))
            $badgeTxt = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#9B2C2C"))
            $g.FillRectangle($badgeBg, $x + 12, $curY + 2, 28, 18)
            $g.DrawString("PK", $fontBadge, $badgeTxt, [float]($x + 17), [float]($curY + 4))
        } elseif ($colKey -eq "FK") {
            $badgeBg = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#FEEBC8"))
            $badgeTxt = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#C05621"))
            $g.FillRectangle($badgeBg, $x + 12, $curY + 2, 28, 18)
            $g.DrawString("FK", $fontBadge, $badgeTxt, [float]($x + 17), [float]($curY + 4))
        } elseif ($colKey -eq "UQ") {
            $badgeBg = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#EBF8FF"))
            $badgeTxt = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#2B6CB0"))
            $g.FillRectangle($badgeBg, $x + 12, $curY + 2, 28, 18)
            $g.DrawString("UQ", $fontBadge, $badgeTxt, [float]($x + 16), [float]($curY + 4))
        } else {
            $dotBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#CBD5E0"))
            $g.FillEllipse($dotBrush, $x + 22, $curY + 8, 6, 6)
        }
        
        # Column Name & Type
        $g.DrawString($colName, $fontColName, $brushCharcoal, [float]($x + 48), [float]($curY + 2))
        $g.DrawString($colType, $fontColType, $brushMuted, [float]($x + $w - 85), [float]($curY + 3))
        
        $curY += $rowHeight
    }
}

# Entity 1: dim_stores (Top Left)
$colsStores = @(
    @("store_id", "TEXT", "PK"),
    @("store_name", "TEXT", ""),
    @("region_id", "TEXT", ""),
    @("region_name", "TEXT", ""),
    @("city", "TEXT", ""),
    @("state", "TEXT", ""),
    @("store_tier", "TEXT", ""),
    @("store_size_sqft", "INTEGER", ""),
    @("opening_date", "TEXT", ""),
    @("is_active", "INTEGER", "")
)
Draw-TableEntity 70 160 380 "dim_stores" "Master Dimension" $colsStores

# Entity 2: dim_sales_targets (Bottom Left)
$colsTargets = @(
    @("target_id", "TEXT", "PK"),
    @("store_id", "TEXT", "FK"),
    @("target_month", "TEXT", ""),
    @("target_revenue_inr", "REAL", ""),
    @("target_orders_count", "INTEGER", ""),
    @("target_atv_inr", "REAL", "")
)
Draw-TableEntity 70 540 380 "dim_sales_targets" "Budget Fact/Dim" $colsTargets

# Entity 3: fact_orders (Center Left)
$colsOrders = @(
    @("order_id", "TEXT", "PK"),
    @("store_id", "TEXT", "FK"),
    @("order_datetime", "TEXT", ""),
    @("order_date", "TEXT", ""),
    @("order_time", "TEXT", ""),
    @("order_hour", "INTEGER", ""),
    @("day_of_week", "TEXT", ""),
    @("is_weekend_flag", "INTEGER", ""),
    @("payment_method", "TEXT", ""),
    @("order_status", "TEXT", ""),
    @("gross_amount_inr", "REAL", ""),
    @("discount_amount_inr", "REAL", ""),
    @("return_amount_inr", "REAL", ""),
    @("net_amount_inr", "REAL", "")
)
Draw-TableEntity 560 160 410 "fact_orders" "Transaction Header Fact" $colsOrders

# Entity 4: fact_order_items (Center Right)
$colsItems = @(
    @("order_item_id", "INTEGER", "PK"),
    @("order_id", "TEXT", "FK"),
    @("product_id", "TEXT", "FK"),
    @("quantity", "INTEGER", ""),
    @("unit_cost_price_inr", "REAL", ""),
    @("unit_selling_price_inr", "REAL", ""),
    @("discount_pct", "REAL", ""),
    @("line_gross_amount_inr", "REAL", ""),
    @("line_discount_amount_inr", "REAL", ""),
    @("line_net_amount_inr", "REAL", ""),
    @("line_cogs_inr", "REAL", ""),
    @("line_profit_inr", "REAL", ""),
    @("line_margin_pct", "REAL", "")
)
Draw-TableEntity 1070 160 410 "fact_order_items" "Line Item Fact" $colsItems

# Entity 5: dim_products (Far Right)
$colsProducts = @(
    @("product_id", "TEXT", "PK"),
    @("product_name", "TEXT", ""),
    @("category_id", "TEXT", ""),
    @("category_name", "TEXT", ""),
    @("sub_category", "TEXT", ""),
    @("cost_price_inr", "REAL", ""),
    @("base_selling_price_inr", "REAL", ""),
    @("target_margin_pct", "REAL", ""),
    @("sku_code", "TEXT", "UQ"),
    @("is_active", "INTEGER", "")
)
Draw-TableEntity 1540 160 340 "dim_products" "Master Dimension" $colsProducts

# Entity 6: view_store_sales_analytics (Bottom Banner View)
$viewBoxX = 560
$viewBoxY = 620
$viewBoxW = 1320
$viewBoxH = 140

$brushViewBg = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#F0FFF4"))
$penViewBorder = New-Object System.Drawing.Pen([System.Drawing.ColorTranslator]::FromHtml("#38A169"), [float]1.8)
$g.FillRectangle($brushViewBg, $viewBoxX, $viewBoxY, $viewBoxW, $viewBoxH)
$g.DrawRectangle($penViewBorder, $viewBoxX, $viewBoxY, $viewBoxW, $viewBoxH)

$fontViewTitle = New-Object System.Drawing.Font("Segoe UI", [float]12, [System.Drawing.FontStyle]::Bold)
$fontViewText = New-Object System.Drawing.Font("Segoe UI", [float]9.5, [System.Drawing.FontStyle]::Regular)
$brushViewTitle = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#22543D"))

$g.DrawString("DENORMALIZED ANALYTICAL VIEW: view_store_sales_analytics (36 Calculated & Dimension Columns)", $fontViewTitle, $brushViewTitle, [float]($viewBoxX + 16), [float]($viewBoxY + 12))
$viewDesc = "Consolidates: order_item_id, order_id, datetime, date, hour, year, month, quarter, day_of_week, weekend_flag, store_id, store_name, region, city, state, store_tier, store_size_sqft, product_id, product_name, category_name, sub_category, quantity, unit_cost, unit_price, discount_pct, line_gross, line_discount, line_net, line_cogs, line_profit, line_margin_pct, payment_method, order_status, sales_velocity_sqft_contrib.`nDesigned for sub-millisecond Streamlit Cloud dashboards with zero runtime multi-table JOIN overhead."
$g.DrawString($viewDesc, $fontViewText, $brushCharcoal, [float]($viewBoxX + 16), [float]($viewBoxY + 40))

# ----------------------------------------------------
# Draw Relationship Lines & Labels
# ----------------------------------------------------
function Draw-RelationLine {
    param(
        [int]$x1, [int]$y1, [int]$x2, [int]$y2, [string]$verb, [string]$cardLeft, [string]$cardRight
    )
    $g.DrawLine($penLine, $x1, $y1, $x2, $y2)
    
    # Circle at parent (1)
    $brushParent = New-Object System.Drawing.SolidBrush($primaryColor)
    $g.FillEllipse($brushParent, $x1 - 5, $y1 - 5, 10, 10)
    
    # Arrow / crow's foot indicator at child (N)
    $g.DrawLine($penLine, $x2, $y2, $x2 - 10, $y2 - 6)
    $g.DrawLine($penLine, $x2, $y2, $x2 - 10, $y2 + 6)
    
    # Verb label in middle
    $midX = ($x1 + $x2) / 2
    $midY = ($y1 + $y2) / 2
    
    $tagBg = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#EDF2F7"))
    $penTag = New-Object System.Drawing.Pen([System.Drawing.ColorTranslator]::FromHtml("#CBD5E0"), [float]1.0)
    $tagSize = $g.MeasureString($verb, $fontVerb)
    $g.FillRectangle($tagBg, [float]($midX - ($tagSize.Width / 2) - 4), [float]($midY - ($tagSize.Height / 2) - 2), [float]($tagSize.Width + 8), [float]($tagSize.Height + 4))
    $g.DrawRectangle($penTag, [float]($midX - ($tagSize.Width / 2) - 4), [float]($midY - ($tagSize.Height / 2) - 2), [float]($tagSize.Width + 8), [float]($tagSize.Height + 4))
    $brushVerb = New-Object System.Drawing.SolidBrush($secondaryColor)
    $g.DrawString($verb, $fontVerb, $brushVerb, [float]($midX - ($tagSize.Width / 2)), [float]($midY - ($tagSize.Height / 2)))
    
    # Cardinality tags
    $fontCard = New-Object System.Drawing.Font("Segoe UI", [float]9, [System.Drawing.FontStyle]::Bold)
    $g.DrawString($cardLeft, $fontCard, $brushPrimary, [float]($x1 + 8), [float]($y1 - 18))
    $g.DrawString($cardRight, $fontCard, $brushPrimary, [float]($x2 - 32), [float]($y2 - 18))
}

# Relation 1: dim_stores -> fact_orders (1:N)
Draw-RelationLine 450 240 560 240 "operates / hosts" "1" "0..N"

# Relation 2: fact_orders -> fact_order_items (1:N)
Draw-RelationLine 970 240 1070 240 "contains" "1" "1..N"

# Relation 3: dim_products -> fact_order_items (1:N)
Draw-RelationLine 1540 240 1480 240 "referenced_in" "1" "0..N"

# Relation 4: dim_stores -> dim_sales_targets (1:N vertical line)
$g.DrawLine($penLine, 260, 480, 260, 540)
$g.FillEllipse($brushPrimary, 255, 475, 10, 10)
$g.DrawLine($penLine, 260, 540, 254, 530)
$g.DrawLine($penLine, 260, 540, 266, 530)

$tagBg = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#EDF2F7"))
$penTag = New-Object System.Drawing.Pen([System.Drawing.ColorTranslator]::FromHtml("#CBD5E0"), [float]1.0)
$tagSize = $g.MeasureString("assigned", $fontVerb)
$g.FillRectangle($tagBg, [float](260 - ($tagSize.Width / 2) - 4), [float](510 - ($tagSize.Height / 2) - 2), [float]($tagSize.Width + 8), [float]($tagSize.Height + 4))
$g.DrawRectangle($penTag, [float](260 - ($tagSize.Width / 2) - 4), [float](510 - ($tagSize.Height / 2) - 2), [float]($tagSize.Width + 8), [float]($tagSize.Height + 4))
$brushVerb = New-Object System.Drawing.SolidBrush($secondaryColor)
$g.DrawString("assigned", $fontVerb, $brushVerb, [float](260 - ($tagSize.Width / 2)), [float](510 - ($tagSize.Height / 2)))

$fontCard = New-Object System.Drawing.Font("Segoe UI", [float]9, [System.Drawing.FontStyle]::Bold)
$g.DrawString("1", $fontCard, $brushPrimary, [float]270, [float]482)
$g.DrawString("0..N", $fontCard, $brushPrimary, [float]270, [float]520)

# Legend Box
$legX = 70
$legY = 820
$legW = 1820
$legH = 100
$brushLegBg = New-Object System.Drawing.SolidBrush($bannerBg)
$g.FillRectangle($brushLegBg, $legX, $legY, $legW, $legH)
$g.DrawRectangle($penBorder, $legX, $legY, $legW, $legH)

$fontLegHdr = New-Object System.Drawing.Font("Segoe UI", [float]10.5, [System.Drawing.FontStyle]::Bold)
$fontLegTxt = New-Object System.Drawing.Font("Segoe UI", [float]9.5, [System.Drawing.FontStyle]::Regular)
$g.DrawString("LEGEND & ARCHITECTURAL CONSTRAINTS:", $fontLegHdr, $brushPrimary, [float]($legX + 20), [float]($legY + 12))

$badgeBgPK = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#FED7D7"))
$badgeTxtPK = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#9B2C2C"))
$g.FillRectangle($badgeBgPK, $legX + 20, $legY + 40, 28, 18)
$g.DrawString("PK", $fontBadge, $badgeTxtPK, [float]($legX + 25), [float]($legY + 42))
$g.DrawString("Primary Key (Unique, Not Null)", $fontLegTxt, $brushCharcoal, [float]($legX + 54), [float]($legY + 40))

$badgeBgFK = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#FEEBC8"))
$badgeTxtFK = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#C05621"))
$g.FillRectangle($badgeBgFK, $legX + 320, $legY + 40, 28, 18)
$g.DrawString("FK", $fontBadge, $badgeTxtFK, [float]($legX + 325), [float]($legY + 42))
$g.DrawString("Foreign Key (Referential Constraint)", $fontLegTxt, $brushCharcoal, [float]($legX + 354), [float]($legY + 40))

$badgeBgUQ = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#EBF8FF"))
$badgeTxtUQ = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#2B6CB0"))
$g.FillRectangle($badgeBgUQ, $legX + 640, $legY + 40, 28, 18)
$g.DrawString("UQ", $fontBadge, $badgeTxtUQ, [float]($legX + 644), [float]($legY + 42))
$g.DrawString("Unique Constraint (Zero Duplicates)", $fontLegTxt, $brushCharcoal, [float]($legX + 674), [float]($legY + 40))

$g.DrawString("Cardinality: 1 to 0..N (One-to-Many Optional) | 1 to 1..N (One-to-Many Mandatory) | Enforced via SQLite Foreign Keys & PRAGMA foreign_keys = ON", $fontLegTxt, $brushMuted, [float]($legX + 20), [float]($legY + 68))

$targetDir = "docs"
if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
$targetPath = Join-Path $targetDir "09_ER_Diagram.png"
$bmp.Save($targetPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
$g.Dispose()

Write-Host "ER Diagram image successfully rendered at: $targetPath"
