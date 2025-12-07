import Handsontable from 'handsontable';
import 'handsontable/dist/handsontable.full.css';
import { HyperFormula } from 'hyperformula';

function SheetEditor({ data, onChange, canEdit }) {
  const containerRef = React.useRef(null);
  const hotRef = React.useRef(null);

  useEffect(() => {
    const hfInstance = HyperFormula.buildEmpty();
    hotRef.current = new Handsontable(containerRef.current, {
      data,
      licenseKey: 'non-commercial-and-evaluation',
      rowHeaders: true,
      colHeaders: true,
      formulas: { engine: hfInstance },
      readOnly: !canEdit,
      contextMenu: canEdit,
      dropdownMenu: true,
      manualRowMove: true,
      manualColumnMove: true,
      manualColumnResize: true,
      afterChange: (changes, source) => {
        if (!changes || source === 'loadData') return;
        const next = hotRef.current.getData();
        onChange(next);
      },
    });
    return () => hotRef.current?.destroy();
  }, []);

  useEffect(() => {
    if (hotRef.current) hotRef.current.loadData(data);
  }, [data]);

  return <div className="hot-container" ref={containerRef} />;
}
