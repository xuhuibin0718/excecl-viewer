import { useEffect, useRef, useState } from 'react';
import PropTypes from 'prop-types';
import * as XLSX from 'xlsx';
import Handsontable from 'handsontable';
import 'handsontable/dist/handsontable.full.min.css';
import './ExcelPreview.css';

const ExcelPreview = ({ fileStream, fileName, onClose }) => {
  const hotContainerRef = useRef(null);
  const hotInstance = useRef(null);
  const [excelData, setExcelData] = useState({
    data: [],
    merges: [],
    styles: []
  });

  useEffect(() => {
    if (fileStream) {
      const workbook = XLSX.read(fileStream, {
        type: "array",
        cellStyles: true,
        codepage: 936,
        dense: true,
        bookVBA: true,
      });

      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        cellStyles: true,
        raw: false,
        defval: "",
      });
      const merges = worksheet["!merges"] || [];

      // Convert styles to Handsontable format
      const styles = worksheet.map((row) =>
        row.map((cell) => {
          if (!cell || !cell.s) return null;
          
          const style = {
            backgroundColor: cell.s?.fgColor?.rgb ? 
              `#${cell.s.fgColor.rgb}` : undefined,
            color: cell.s?.fgColor?.rgb ? 
              '#000000' : undefined,
            fontWeight: cell.s?.bold ? 'bold' : undefined,
            fontStyle: cell.s?.italic ? 'italic' : undefined,
            fontSize: cell.s?.sz ? `${cell.s.font.sz}pt` : undefined,
            fontFamily: cell.s?.name || undefined,
            textAlign: cell.s.alignment?.horizontal || 'center',
            verticalAlign: cell.s.alignment?.vertical || 'middle',
            border: cell.s.border ? {
              top: cell.s.border.top ? {
                width: cell.s.border.top.style || 1,
                color: cell.s.border.top.color?.rgb ? 
                  `#${cell.s.border.top.color.rgb}` : '#FFFFFF'
              } : undefined,
              right: cell.s.border.right ? {
                width: cell.s.border.right.style || 1,
                color: cell.s.border.right.color?.rgb ? 
                  `#${cell.s.border.right.color.rgb}` : '#FFFFFF'
              } : undefined,
              bottom: cell.s.border.bottom ? {
                width: cell.s.border.bottom.style || 1,
                color: cell.s.border.bottom.color?.rgb ? 
                  `#${cell.s.border.bottom.color.rgb}` : '#FFFFFF'
              } : undefined,
              left: cell.s.border.left ? {
                width: cell.s.border.left.style || 1,
                color: cell.s.border.left.color?.rgb ? 
                  `#${cell.s.border.left.color.rgb}` : '#FFFFFF'
              } : undefined
            } : undefined,
            format: cell.s.numFmt || undefined
          };

          // Remove undefined properties
          Object.keys(style).forEach(key => {
            if (style[key] === undefined) {
              delete style[key];
            }
          });

          return Object.keys(style).length > 0 ? style : null;
        })
      );

      setExcelData({
        data,
        merges: merges.map((m) => ({
          row: m.s.r,
          col: m.s.c,
          rowspan: m.e.r - m.s.r + 1,
          colspan: m.e.c - m.s.c + 1,
        })),
        styles
      });
    }
  }, [fileStream]);

  useEffect(() => {
    if (hotContainerRef.current && excelData.data.length > 0) {
      hotInstance.current = new Handsontable(hotContainerRef.current, {
        data: excelData.data,
        colHeaders: false,
        rowHeaders: false,
        mergeCells: excelData.merges,
        autoRowSize: true,
        autoColumnSize: true,
        manualRowResize: true,
        manualColumnResize: true,
        stretchH: 'all',
        height: '100%',
        width: '100%',
        wordWrap: true,
        cells: (row, col) => {
          const cellProperties = {};
          
          cellProperties.renderer = function(instance, td) {
            Handsontable.renderers.TextRenderer.apply(this, arguments);
            
            if (excelData.styles && excelData.styles[row] && excelData.styles[row][col]) {
              const style = excelData.styles[row][col];
              
              Object.entries(style).forEach(([key, value]) => {
                if (key === 'border') {
                  Object.entries(value).forEach(([side, borderStyle]) => {
                    if (borderStyle) {
                      td.style[`border${side.charAt(0).toUpperCase() + side.slice(1)}`] = 
                        `${borderStyle.width}px solid ${borderStyle.color}`;
                    }
                  });
                } else if (key === 'format') {
                  const cellValue = instance.getDataAtCell(row, col);
                  if (typeof cellValue === 'number') {
                    if (value === '0.00%' || value === '0%') {
                      td.textContent = (cellValue * 100).toFixed(2) + '%';
                    } else {
                      td.textContent = cellValue.toFixed(2);
                    }
                  }
                } else {
                  td.style[key] = value;
                }
              });
            }
            
            if (!td.style.backgroundColor) {
              if (row < 2) {
                td.style.backgroundColor = 'rgba(23, 50, 109, 0.66)';
                td.style.color = 'white';
              } else {
                td.style.backgroundColor = 'rgba(10, 42, 85, 1)';
                td.style.color = 'white';
              }
            }
            
            td.style.textAlign = 'center';
            td.style.verticalAlign = 'middle';
            td.style.whiteSpace = 'normal';
            td.style.wordBreak = 'break-word';
            td.style.maxHeight = '2.4em';
            td.style.overflow = 'hidden';
            td.style.lineHeight = '1.2em';
            
            const cellValue = instance.getDataAtCell(row, col);
            if (cellValue && typeof cellValue === 'string') {
              const lines = cellValue.split('\n');
              if (lines.length > 2) {
                td.textContent = lines.slice(0, 2).join('\n') + '...';
              }
            }
          };
          
          return cellProperties;
        },
        afterRender: function() {
          this.getPlugin('autoColumnSize').recalculateAllColumnsWidth();
        }
      });
    }

    return () => {
      if (hotInstance.current) {
        hotInstance.current.destroy();
      }
    };
  }, [excelData]);

  return (
    <div className="excel-preview-modal">
      <div className="excel-preview-content">
        <div className="close-button" onClick={onClose}>×</div>
        <div className="excel-preview-header">
          <div className="excel-preview-title">{fileName}</div>
        </div>
        <div className="excel-preview-body">
          <div className="handsontable-container" ref={hotContainerRef}></div>
        </div>
      </div>
    </div>
  );
};

ExcelPreview.propTypes = {
  fileStream: PropTypes.instanceOf(ArrayBuffer).isRequired,
  fileName: PropTypes.string.isRequired,
  onClose: PropTypes.func.isRequired
};

export default ExcelPreview; 