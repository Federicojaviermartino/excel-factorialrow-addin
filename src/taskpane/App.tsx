import { useEffect, useState, useCallback } from "react";

type Orientation = "row" | "column";

async function readOrientation(): Promise<Orientation> {
  try {
    const value = await OfficeRuntime.storage.getItem("orientation");
    return value === "column" ? "column" : "row";
  } catch (error) {
    console.warn('Failed to read orientation from storage:', error);
    return "row";
  }
}

async function writeOrientation(v: Orientation): Promise<void> {
  try {
    await OfficeRuntime.storage.setItem("orientation", v);
  } catch (error) {
    console.error('Failed to write orientation to storage:', error);
    throw error;
  }
}

export function App() {
  const [orientation, setOrientation] = useState<Orientation>("row");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    readOrientation().then(setOrientation).catch(err => {
      console.error('Error loading initial orientation:', err);
      setError('Failed to load settings');
    });
  }, []);

  const onChange = useCallback(async (v: Orientation) => {
    setBusy(true);
    setError(null);
    
    try {
      setOrientation(v);
      await writeOrientation(v);
      
      await Excel.run(async (ctx) => {
        ctx.workbook.application.calculate(Excel.CalculationType.full);
        await ctx.sync();
      });
    } catch (error) {
      console.error('Failed to update orientation:', error);
      setError('Failed to update settings. Please try again.');
      const currentOrientation = await readOrientation();
      setOrientation(currentOrientation);
    } finally {
      setBusy(false);
    }
  }, []);

  return (
    <div style={{ fontFamily: 'system-ui, Arial, sans-serif', padding: 16, lineHeight: 1.4 }}>
      <h2 style={{ marginTop: 0 }}>Factorial: Orientation</h2>
      <p>Select how the results of <code>TESTVELIXO.FACTORIALROW</code> should spill.</p>
      
      {error && (
        <div style={{ 
          backgroundColor: '#fee', 
          border: '1px solid #fcc', 
          borderRadius: 4, 
          padding: 8, 
          marginBottom: 16,
          fontSize: 14,
          color: '#c44'
        }}>
          ⚠️ {error}
        </div>
      )}
      
      <div role="radiogroup" aria-label="Orientation">
        <label style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 8 }}>
          <input
            type="radio"
            name="orientation"
            checked={orientation === "row"}
            onChange={() => onChange("row")}
            disabled={busy}
          />
          <span>Row</span>
        </label>
        <label style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <input
            type="radio"
            name="orientation"
            checked={orientation === "column"}
            onChange={() => onChange("column")}
            disabled={busy}
          />
          <span>Column</span>
        </label>
      </div>
      
      <p style={{ fontSize: 12, color: '#555' }}>
        The preference is saved and shared with custom functions via <code>OfficeRuntime.storage</code>.
      </p>
      
      <hr />
      <section>
        <h3>Usage</h3>
        <ol>
          <li>In a cell, type: <code>=TESTVELIXO.FACTORIALROW(10)</code></li>
          <li>Use the buttons above to toggle between <strong>Row</strong> and <strong>Column</strong>.</li>
        </ol>
      </section>
    </div>
  );
}
