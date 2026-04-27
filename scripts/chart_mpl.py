#!/usr/bin/env python3
"""
matplotlib-based chart generator for multi-series line charts.
Designed for the weekly report - readable, publication-quality output.

Usage:
  python3 chart_mpl.py --data '[{"x":"2026-01-01","y":5.2,"market":"Brent"}]' \\
    --output chart.png --title "My Chart"
"""
import argparse, json, sys, os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mticker
from datetime import datetime
import numpy as np

# ── Configuration ──
FONT_FAMILY = 'Helvetica, Arial, sans-serif'
CHART_DPI = 120
CHART_WIDTH = 16   # inches
CHART_HEIGHT = 8   # inches

# Differentiate up to 10 series with color + line style
COLORS = ['#e74c3c', '#2980b9', '#27ae60', '#8e44ad', '#f39c12', 
          '#1abc9c', '#e67e22', '#2c3e50', '#c0392b', '#16a085']
LINE_STYLES = ['-', '--', '-.', ':', (0, (3, 1, 1, 1)), (0, (5, 2)),
               (0, (1, 1)), (0, (3, 1, 2, 1, 1, 1)), '-', '--']

def parse_args():
    parser = argparse.ArgumentParser(description='Generate multi-series line chart')
    parser.add_argument('--data', help='JSON array of data points')
    parser.add_argument('--output', default='chart.png', help='Output PNG path')
    parser.add_argument('--title', default='', help='Chart title')
    parser.add_argument('--subtitle', default='', help='Chart subtitle')
    parser.add_argument('--x-field', default='x', help='X axis field')
    parser.add_argument('--y-field', default='y', help='Y axis field')
    parser.add_argument('--series-field', default='market', help='Series split field')
    parser.add_argument('--x-title', default='Date', help='X axis label')
    parser.add_argument('--y-title', default='% Change', help='Y axis label')
    parser.add_argument('--y-format', default='.1f', help='Y axis number format')
    parser.add_argument('--width', type=int, default=CHART_WIDTH)
    parser.add_argument('--height', type=int, default=CHART_HEIGHT)
    parser.add_argument('--dpi', type=int, default=CHART_DPI)
    parser.add_argument('--x-type', default='temporal', help='temporal | ordinal')
    parser.add_argument('--x-format', default='%b %d', help='Date format for temporal x')
    parser.add_argument('--hline', default='', help='Horizontal line as "value,color"')
    parser.add_argument('--last-value', action='store_true', help='Label last value')
    parser.add_argument('--y-domain', default='', help='Y axis min,max (e.g., "0,100")')
    parser.add_argument('--auto-focus-y', action='store_true', help='Auto focus Y on data range')
    parser.add_argument('--legend-columns', type=int, default=0, help='Number of legend columns')
    return parser.parse_args()

def main():
    args = parse_args()
    
    # Parse data
    data = json.loads(args.data)
    if not data:
        print("ERROR: Empty data")
        sys.exit(1)
    
    # Group by series
    series = {}
    for pt in data:
        key = pt.get(args.series_field, 'default')
        if key not in series:
            series[key] = {'x': [], 'y': []}
        series[key]['x'].append(pt[args.x_field])
        series[key]['y'].append(pt[args.y_field])
    
    # Create figure
    fig, ax = plt.subplots(figsize=(args.width, args.height), dpi=args.dpi)
    fig.patch.set_facecolor('white')
    
    is_temporal = args.x_type == 'temporal'
    
    for i, (name, sdata) in enumerate(sorted(series.items())):
        x_vals = sdata['x']
        y_vals = sdata['y']
        
        # Convert dates if temporal
        if is_temporal:
            x_vals = [datetime.strptime(str(d)[:10], '%Y-%m-%d') for d in x_vals]
        
        color = COLORS[i % len(COLORS)]
        ls = LINE_STYLES[i % len(LINE_STYLES)]
        
        ax.plot(x_vals, y_vals, color=color, linestyle=ls, linewidth=2.5,
                label=name, alpha=0.85, marker='', markersize=0)
        
        # Last value label
        if args.last_value and y_vals:
            last_y = y_vals[-1]
            last_x = x_vals[-1] if is_temporal else x_vals[-1]
            sign = '+' if last_y >= 0 else ''
            # Format based on value magnitude
            if abs(last_y) < 10:
                label = f'{sign}{last_y:.2f}'
            elif abs(last_y) < 100:
                label = f'{sign}{last_y:.1f}'
            else:
                label = f'{sign}{last_y:.0f}'
            ax.annotate(label,
                        xy=(last_x, last_y),
                        xytext=(8, 0), textcoords='offset points',
                        fontsize=12, fontweight='bold', color=color,
                        va='center')
    
    # Horizontal zero line
    ax.axhline(y=0, color='#999999', linewidth=1, linestyle='-', alpha=0.6)
    
    # Extra hline if specified
    if args.hline:
        parts = args.hline.split(',')
        h_val = float(parts[0])
        h_color = parts[1] if len(parts) > 1 else '#999999'
        ax.axhline(y=h_val, color=h_color, linewidth=0.8, linestyle='--', alpha=0.4)
    
    # Auto-focus Y on data range with 10% padding on each side
    if args.auto_focus_y:
        all_y = [d[args.y_field] for d in data]
        y_min, y_max = min(all_y), max(all_y)
        y_range = y_max - y_min
        padding = y_range * 0.1 if y_range > 0 else y_range * 1.0
        ax.set_ylim(y_min - padding, y_max + padding)
    
    # Explicit y-domain
    if args.y_domain:
        parts = args.y_domain.split(',')
        ax.set_ylim(float(parts[0]), float(parts[1]))
    
    # Grid
    ax.grid(True, alpha=0.25, linestyle='-', linewidth=0.5)
    
    # X-axis for temporal data
    if is_temporal and len(data) > 1:
        # Auto-detect time span
        dates_list = sorted(set(
            datetime.strptime(str(d[args.x_field])[:10], '%Y-%m-%d') 
            for d in data
        ))
        if len(dates_list) <= 1:
            pass  # fallback
        else:
            span_days = (dates_list[-1] - dates_list[0]).days
            
            if span_days <= 30:
                # < 1 month: weekly
                ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%b-%Y'))
            elif span_days <= 90:
                # 1-3 months: biweekly
                ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=2))
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%b-%Y'))
            elif span_days <= 180:
                # 3-6 months: monthly
                ax.xaxis.set_major_locator(mdates.MonthLocator(interval=1))
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%b-%Y'))
                ax.xaxis.set_minor_locator(mdates.WeekdayLocator(interval=2))
            else:
                # >6 months: quarterly
                ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%b-%Y'))
                ax.xaxis.set_minor_locator(mdates.MonthLocator(interval=1))
    
    # Font sizes
    TITLE_FONT = 20
    AXIS_LABEL_FONT = 16
    TICK_FONT = 14
    LEGEND_FONT = 14
    
    ax.set_xlabel(args.x_title, fontsize=AXIS_LABEL_FONT, fontweight='bold', labelpad=8)
    ax.set_ylabel(args.y_title, fontsize=AXIS_LABEL_FONT, fontweight='bold', labelpad=8)
    ax.tick_params(axis='both', labelsize=TICK_FONT)
    ax.tick_params(axis='x', rotation=0)
    
    # Use scientific/compact y-axis format
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter(f'%{args.y_format}'))
    
    # Title + subtitle
    if args.title:
        title_text = args.title
        if args.subtitle:
            title_text += '\n' + args.subtitle
        ax.set_title(title_text, fontsize=TITLE_FONT, fontweight='bold', pad=15, linespacing=1.3)
    
    # Legend
    n_series = len(series)
    if args.legend_columns > 0:
        ncols = args.legend_columns
    elif n_series <= 4:
        ncols = n_series
    elif n_series <= 6:
        ncols = 3
    elif n_series <= 8:
        ncols = 4
    else:
        ncols = n_series // 2
    
    legend = ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.10),
                       ncol=ncols, fontsize=LEGEND_FONT, frameon=True,
                       fancybox=True, shadow=False, edgecolor='#cccccc',
                       handlelength=2.5)
    
    # Tight layout with room for legend below
    plt.tight_layout(rect=[0, 0.05, 1, 0.98])
    
    # Save
    os.makedirs(os.path.dirname(args.output) or '.', exist_ok=True)
    plt.savefig(args.output, bbox_inches='tight', dpi=args.dpi, 
                facecolor='white', edgecolor='none')
    plt.close()
    
    print(f"✅ Chart saved: {args.output} ({os.path.getsize(args.output)/1024:.0f} KB)")


if __name__ == '__main__':
    main()
