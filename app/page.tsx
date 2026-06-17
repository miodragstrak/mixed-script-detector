"use client";

import { useState } from "react";
import * as XLSX from "xlsx";

const CYRILLIC_RE =
  /[А-Яа-яЉЊЏЂЋЖЧШљњџђћжчш]/;

const LATIN_RE =
  /[A-Za-zČĆŽŠĐčćžšđ]/;

function hasMixedScript(text: string): boolean {
  return (
    CYRILLIC_RE.test(text) &&
    LATIN_RE.test(text)
  );
}

// ⬇️ NOVA FUNKCIJA IDE OVDE

function analyzeScripts(text: string) {
  const latin: string[] = [];
  const cyrillic: string[] = [];

  for (const char of text) {
    const code = `U+${char
      .codePointAt(0)!
      .toString(16)
      .toUpperCase()
      .padStart(4, "0")}`;

    if (LATIN_RE.test(char)) {
      latin.push(`${char} [${code}]`);
    }

    if (CYRILLIC_RE.test(char)) {
      cyrillic.push(`${char} [${code}]`);
    }
  }

  return {
    latin: latin.join(", "),
    cyrillic: cyrillic.join(", "),
  };
}

function exportToCsv(
  fileName: string,
  data: {
    excelRow: number;
    column: string;
    value: string;
    latin: string;
    cyrillic: string;
  }[]
) {
  const header = [
    "Excel red",
    "Kolona",
    "Vrednost",
    "Latinični Unicode",
    "Ćirilični Unicode",
  ];

  const rows = data.map((row) => [
    row.excelRow,
    row.column,
    row.value,
    row.latin,
    row.cyrillic,
  ]);

  const csv = [header, ...rows]
    .map((row) =>
      row
        .map((value) => `"${String(value).replace(/"/g, '""')}"`)
        .join(",")
    )
    .join("\n");

  // UTF-8 BOM da Excel pravilno otvori ćirilicu
  const blob = new Blob(["\uFEFF" + csv], {
    type: "text/csv;charset=utf-8;",
  });

  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;

  const baseName = fileName.replace(/\.[^.]+$/, "");

  link.download = `${baseName}_mixed_script_report.csv`;

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
}

export default function Home() {
  const [fileName, setFileName] = useState("");
  const [rowCount, setRowCount] = useState(0);
  const [problemRows, setProblemRows] = useState(0);

  const [problemDetails, setProblemDetails] = useState<
    {
      excelRow: number;
      column: string;
      value: string;
      latin: string;
      cyrillic: string;
    }[]
  >([]);

  const handleFileChange = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0];

    if (!file) return;
    setRowCount(0);
    setProblemRows(0);
    setProblemDetails([]);

    setFileName(file.name);

    const buffer = await file.arrayBuffer();

    const workbook = XLSX.read(buffer);

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const data = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet);

      setRowCount(data.length);

      let mixedCount = 0;

      const problems: {
        excelRow: number;
        column: string;
        value: string;
        latin: string;
        cyrillic: string;
      }[] = [];

      for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
        const row = data[rowIndex];

        let rowHasProblem = false;

        for (const [column, value] of Object.entries(row)) {
          if (typeof value !== "string") continue;

          if (hasMixedScript(value)) {
            rowHasProblem = true;

            const scripts = analyzeScripts(value);

            problems.push({
              excelRow: rowIndex + 2,
              column,
              value,
              latin: scripts.latin,
              cyrillic: scripts.cyrillic,
            });

            break;
          }
        }

        if (rowHasProblem) {
          mixedCount++;
        }
      }

      setProblemRows(mixedCount);
      setProblemDetails(problems);

      console.log("Problematic rows:", mixedCount);
  };

  return (
    <main className="min-h-screen bg-slate-50">
      <div className="max-w-6xl mx-auto px-8 py-12">
        <h1 className="text-4xl font-bold text-slate-900">
          Detektor mešanog pisma
        </h1>

        <p className="mt-3 text-slate-600">
          Pronalaženje zapisa koji sadrže mešovitu upotrebu
          ćirilice i latinice.
        </p>

        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-8 mt-8">
          <label className="block cursor-pointer">
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-8 mt-8 text-center hover:bg-slate-50 transition">
              <div className="text-4xl mb-3">
                📄
              </div>

              <div className="text-lg font-semibold text-slate-900">
                Izaberite Excel fajl
              </div>

              <div className="text-slate-500 mt-2">
                Upload novog fajla automatski pokreće novu analizu
              </div>
            </div>

            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={handleFileChange}
              className="hidden"
            />
          </label>
        </div>

        {fileName && (
          <div className="mt-4 bg-blue-50 border border-blue-200 rounded-xl p-4">
            <div className="text-sm text-blue-600 font-medium">
              Učitani fajl
            </div>

            <div className="text-lg font-semibold text-slate-900">
              📄 {fileName}
            </div>
          </div>
        )}

        {rowCount > 0 && (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-8">

            <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
              <div className="text-sm text-slate-500 uppercase tracking-wide">
                Ukupno redova
              </div>

              <div className="text-4xl font-bold text-slate-900 mt-2">
                {rowCount.toLocaleString("sr-RS")}
              </div>
            </div>

            <div className="bg-red-50 rounded-xl shadow-sm border border-red-200 p-6">
              <div className="text-sm text-red-600 uppercase tracking-wide">
                Problematični redovi
              </div>

              <div className="text-4xl font-bold text-red-700 mt-2">
                {problemRows.toLocaleString("sr-RS")}
              </div>
            </div>

          </div>
        )}

        {problemDetails.length > 0 && (
          <div className="mt-6">
            <h2 className="text-xl font-semibold mb-2">
              Problematični redovi
            </h2>

            <div className="mb-4">
              <button
                onClick={() => exportToCsv(fileName, problemDetails)}
                className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition"
              >
                ⬇️ Preuzmi rezultate u CSV
              </button>
            </div>

            <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden mt-8">
              <table className="w-full">
              <thead className="bg-slate-100">
                <tr>
                  <th className="px-4 py-3 text-left text-sm font-semibold text-slate-700">
                    Excel red
                  </th>
                  <th className="px-4 py-3 text-left text-sm font-semibold text-slate-700">
                    Kolona
                  </th>
                  <th className="px-4 py-3 text-left text-sm font-semibold text-slate-700">
                    Vrednost
                  </th>
                  <th className="px-4 py-3 text-left text-sm font-semibold text-slate-700">
                    Latinični karakteri
                  </th>
                  <th className="px-4 py-3 text-left text-sm font-semibold text-slate-700">
                    Ćirilični karakteri
                  </th>
                </tr>
              </thead>

              <tbody>
                {problemDetails.map((row, index) => (
                  <tr key={index}>
                    <td className="px-4 py-3 border-t border-slate-100">
                      {row.excelRow}
                    </td>
                    <td className="px-4 py-3 border-t border-slate-100">
                      {row.column}
                    </td>

                    <td className="px-4 py-3 border-t border-slate-100">
                      {row.value}
                    </td>
                    <td className="px-4 py-3 border-t border-slate-100">
                      {row.latin}
                    </td>
                    <td className="px-4 py-3 border-t border-slate-100">
                      {row.cyrillic}
                    </td>
                  </tr>
                ))}
              </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </main>
  );
}