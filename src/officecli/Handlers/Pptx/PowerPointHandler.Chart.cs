// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Chart GraphicFrame Builder (PPTX-specific) ====================

    /// <summary>
    /// Create a GraphicFrame embedding a chart and add it to the slide's shape tree.
    /// </summary>
    private static GraphicFrame BuildChartGraphicFrame(
        SlidePart slidePart, ChartPart chartPart, uint shapeId, string name,
        long x, long y, long cx, long cy)
    {
        var relId = slidePart.GetIdOfPart(chartPart);

        var graphicFrame = new GraphicFrame();
        graphicFrame.NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
            new NonVisualDrawingProperties { Id = shapeId, Name = name },
            new NonVisualGraphicFrameDrawingProperties(),
            new ApplicationNonVisualDrawingProperties()
        );
        graphicFrame.Transform = new Transform(
            new Drawing.Offset { X = x, Y = y },
            new Drawing.Extents { Cx = cx, Cy = cy }
        );

        var chartRef = new C.ChartReference { Id = relId };
        graphicFrame.AppendChild(new Drawing.Graphic(
            new Drawing.GraphicData(chartRef)
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
            }
        ));

        return graphicFrame;
    }

    // ==================== Chart Readback (PPTX-specific: reads position from GraphicFrame) ====================

    /// <summary>
    /// Build a DocumentNode from a chart GraphicFrame.
    /// </summary>
    private static DocumentNode ChartToNode(GraphicFrame gf, SlidePart slidePart, int slideNum, int chartIdx, int depth)
    {
        var name = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Chart";

        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/chart[{chartIdx}]",
            Type = "chart",
            Preview = name
        };

        node.Format["name"] = name;

        // Position (PPTX-specific: from GraphicFrame transform)
        var offset = gf.Transform?.Offset;
        if (offset != null)
        {
            if (offset.X is not null) node.Format["x"] = FormatEmu(offset.X!);
            if (offset.Y is not null) node.Format["y"] = FormatEmu(offset.Y!);
        }
        var extents = gf.Transform?.Extents;
        if (extents != null)
        {
            if (extents.Cx is not null) node.Format["width"] = FormatEmu(extents.Cx!);
            if (extents.Cy is not null) node.Format["height"] = FormatEmu(extents.Cy!);
        }

        // Read chart data from ChartPart (shared logic)
        var chartRef = gf.Descendants<C.ChartReference>().FirstOrDefault();
        if (chartRef?.Id?.Value != null)
        {
            try
            {
                var chartPart = (ChartPart)slidePart.GetPartById(chartRef.Id.Value);
                var chartSpace = chartPart.ChartSpace;
                var chart = chartSpace?.GetFirstChild<C.Chart>();
                if (chart != null)
                    ChartHelper.ReadChartProperties(chart, node, depth);
            }
            catch { }
        }

        return node;
    }
}
