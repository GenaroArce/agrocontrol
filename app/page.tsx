"use client";

import React, { useState } from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { Plus, Trash2, Download } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";

const EXCEL_CONFIG = {
  celdaNumero: 'C1',      
  celdaFecha: 'I1',       
  celdaLaborCultivo: 'A2',   
  filaNombresProductos: 3, 
  filaInicioDatos: 6,
  filaTotales: 24, 
  maxLotes: 9, 
  colCampo: 1,
  colHas: 2,
  colsProductos: [3, 4, 5, 6, 7], 
  colsCoadyuvantes: [8, 9, 10], 
  celdaContratistaValor: 'D26', 
  celdaMaquinaValor: 'D28',     
  celdaObservaciones: 'A31', 
};

interface FilaOrden {
  id: number;
  campo: string;
  has: number;
  dosis: (number | string)[];
  coadyuvantes: (number | string)[];
}

export default function OrdenTrabajoApp() {
  const [datosGrales, setDatosGrales] = useState({
    numero: "",
    fecha: new Date().toISOString().split("T")[0],
    labor: "",
    cultivo: "",
    contratista: "",
    maquina: "",
    observaciones: "",
    coadyuvantesTexto: "",
  });

  const [nombresProductos, setNombresProductos] = useState(["", "", "", "", ""]);
  const [nombresCoadyuvantes, setNombresCoadyuvantes] = useState(["", "", ""]);

  const [filas, setFilas] = useState<FilaOrden[]>([
    { id: 1, campo: "Lt 1", has: 0, dosis: [0, 0, 0, 0, 0], coadyuvantes: [0, 0, 0] },
  ]);

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    setDatosGrales({ ...datosGrales, [e.target.name]: e.target.value });
  };

  const handleProdNameChange = (index: number, val: string) => {
    const nuevos = [...nombresProductos];
    nuevos[index] = val;
    setNombresProductos(nuevos);
  };

  const handleCoadNameChange = (index: number, val: string) => {
    const nuevos = [...nombresCoadyuvantes];
    nuevos[index] = val;
    setNombresCoadyuvantes(nuevos);
  };

  const handleFilaChange = (id: number, field: string, value: string | number, indexArray?: number) => {
    setFilas(filas.map((f) => {
        if (f.id !== id) return f;
        
        if (field === 'dosis' && indexArray !== undefined) {
            const nuevasDosis = [...f.dosis];
            nuevasDosis[indexArray] = value;
            return { ...f, dosis: nuevasDosis };
        }

        if (field === 'coadyuvantes' && indexArray !== undefined) {
            const nuevosCoad = [...f.coadyuvantes];
            nuevosCoad[indexArray] = value;
            return { ...f, coadyuvantes: nuevosCoad };
        }

        return { ...f, [field]: value };
    }));
  };

  const agregarFila = () => {
    if (filas.length >= EXCEL_CONFIG.maxLotes) return;
    setFilas([...filas, { 
        id: Date.now(), 
        campo: `Lt ${filas.length + 1}`, 
        has: 0, 
        dosis: [0, 0, 0, 0, 0],
        coadyuvantes: [0, 0, 0]
    }]);
  };

  const eliminarFila = (id: number) => {
    setFilas(filas.filter((f) => f.id !== id));
  };

  const generarExcel = async () => {
    try {
      const response = await fetch("/plantilla.xlsx");
      if (!response.ok) throw new Error("Falta plantilla.xlsx");
      const buffer = await response.arrayBuffer();
      
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const sheet = workbook.getWorksheet(1);
      if (!sheet) return;

      const tituloCell = sheet.getCell('A1');
      const tituloVal = tituloCell.value ? tituloCell.value.toString() : "ORDEN DE TRABAJO N°:";
      if (!tituloVal.includes(datosGrales.numero)) {
          tituloCell.value = `${tituloVal}  ${datosGrales.numero}`;
      }

      sheet.getCell(EXCEL_CONFIG.celdaFecha).value = new Date(datosGrales.fecha).toLocaleDateString("es-AR");

      const celdaCombinada = sheet.getCell(EXCEL_CONFIG.celdaLaborCultivo);
      celdaCombinada.value = `LABOR: ${datosGrales.labor}\nCULTIVO: ${datosGrales.cultivo}`;
      celdaCombinada.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };

      const rowNombres = sheet.getRow(EXCEL_CONFIG.filaNombresProductos);
      nombresProductos.forEach((nombre, idx) => {
          rowNombres.getCell(EXCEL_CONFIG.colsProductos[idx]).value = nombre;
      });
      nombresCoadyuvantes.forEach((nombre, idx) => {
          rowNombres.getCell(EXCEL_CONFIG.colsCoadyuvantes[idx]).value = nombre;
      });

      filas.forEach((fila, index) => {
        const rowWhiteNum = EXCEL_CONFIG.filaInicioDatos + (index * 2); 
        const rowGreyNum = rowWhiteNum + 1; 
        
        const rowWhite = sheet.getRow(rowWhiteNum);
        const rowGrey = sheet.getRow(rowGreyNum);

        rowWhite.getCell(EXCEL_CONFIG.colCampo).value = fila.campo;
        rowWhite.getCell(EXCEL_CONFIG.colHas).value = fila.has;

        const refHas = `B${rowWhiteNum}`; 

        fila.dosis.forEach((val, idx) => {
            const col = EXCEL_CONFIG.colsProductos[idx];
            rowWhite.getCell(col).value = (val === 0 || val === '0' || val === '' || val === null) ? null : Number(val);
            const cellAddr = rowWhite.getCell(col).address;
            rowGrey.getCell(col).value = { formula: `IF(ISNUMBER(${cellAddr}), ${cellAddr}*${refHas}, 0)` };
        });

        fila.coadyuvantes.forEach((val, idx) => {
            const col = EXCEL_CONFIG.colsCoadyuvantes[idx];
            rowWhite.getCell(col).value = (val === 0 || val === '0' || val === '' || val === null) ? null : Number(val);
            const cellAddr = rowWhite.getCell(col).address;
            rowGrey.getCell(col).value = { formula: `IF(ISNUMBER(${cellAddr}), ${cellAddr}*${refHas}, 0)` };
        });

        rowWhite.commit();
        rowGrey.commit();
      });

      const rowTotales = sheet.getRow(EXCEL_CONFIG.filaTotales);

      const filasBlancas = Array.from({length: EXCEL_CONFIG.maxLotes}, (_, i) => EXCEL_CONFIG.filaInicioDatos + (i * 2));
      const formulaHas = filasBlancas.map(r => `B${r}`).join('+');
      rowTotales.getCell(EXCEL_CONFIG.colHas).value = { formula: formulaHas };

      const filasGrises = Array.from({length: EXCEL_CONFIG.maxLotes}, (_, i) => EXCEL_CONFIG.filaInicioDatos + 1 + (i * 2));
      
      for (let col = 3; col <= 10; col++) {
          const letraColumna = String.fromCharCode(64 + col);
          const formulaProd = filasGrises.map(r => `${letraColumna}${r}`).join('+');
          rowTotales.getCell(col).value = { formula: formulaProd };
      }
      
      rowTotales.commit();

      sheet.getCell(EXCEL_CONFIG.celdaContratistaValor).value = datosGrales.contratista;
      sheet.getCell(EXCEL_CONFIG.celdaMaquinaValor).value = datosGrales.maquina;
      sheet.getCell(EXCEL_CONFIG.celdaObservaciones).value = `OBSERVACIONES:\n${datosGrales.observaciones}`;

      const outBuffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([outBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, `Orden_${datosGrales.numero}.xlsx`);

    } catch (error) {
      console.error(error);
      alert("Error con la plantilla.xlsx");
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 font-sans text-slate-800">
      <div className="max-w-7xl mx-auto space-y-6">
        
        <div className="flex justify-between items-center bg-white p-4 rounded-lg shadow-sm border border-slate-200">
            <div>
                <h1 className="text-2xl font-bold text-slate-900">AgroControl</h1>
                <p className="text-xs text-slate-500">Creado por Genaro Arce</p>
            </div>
            <Button className="bg-green-600 hover:bg-green-700 text-white font-bold" onClick={generarExcel}>
                <Download className="mr-2 h-5 w-5" /> Descargar
            </Button>
        </div>

        <Card className="shadow-sm">
            <CardHeader className="pb-3 border-b bg-slate-50/50"><CardTitle className="text-base font-bold text-slate-700">Encabezado</CardTitle></CardHeader>
            <CardContent className="grid grid-cols-2 sm:grid-cols-4 gap-4 pt-4">
                <div className="col-span-1"><Label>N° Orden</Label><Input name="numero" value={datosGrales.numero} onChange={handleInputChange} className="font-bold border-slate-300"/></div>
                <div className="col-span-1"><Label>Fecha</Label><Input type="date" name="fecha" value={datosGrales.fecha} onChange={handleInputChange} /></div>
                <div className="col-span-1"><Label>Labor</Label><Input name="labor" value={datosGrales.labor} onChange={handleInputChange} placeholder="Pulverizada" /></div>
                <div className="col-span-1"><Label>Cultivo</Label><Input name="cultivo" value={datosGrales.cultivo} onChange={handleInputChange} placeholder="Soja" /></div>
                <div className="col-span-2"><Label>Contratista</Label><Input name="contratista" value={datosGrales.contratista} onChange={handleInputChange} /></div>
                <div className="col-span-2"><Label>Máquina</Label><Input name="maquina" value={datosGrales.maquina} onChange={handleInputChange} /></div>
            </CardContent>
        </Card>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <Card className="border-t-4 border-t-blue-500 shadow-sm">
                <CardHeader className="py-2 bg-blue-50"><CardTitle className="text-sm font-bold text-blue-700">Productos</CardTitle></CardHeader>
                <CardContent className="pt-4 grid grid-cols-5 gap-2">
                    {nombresProductos.map((prod, idx) => (
                        <Input key={idx} value={prod} onChange={(e) => handleProdNameChange(idx, e.target.value)} className="text-center text-xs h-8 bg-white" placeholder={`P${idx+1}`}/>
                    ))}
                </CardContent>
            </Card>

            <Card className="border-t-4 border-t-yellow-500 shadow-sm">
                <CardHeader className="py-2 bg-yellow-50"><CardTitle className="text-sm font-bold text-yellow-700">Coadyuvantes</CardTitle></CardHeader>
                <CardContent className="pt-4 grid grid-cols-3 gap-2">
                    {nombresCoadyuvantes.map((coad, idx) => (
                        <Input key={idx} value={coad} onChange={(e) => handleCoadNameChange(idx, e.target.value)} className="text-center text-xs h-8 bg-white" placeholder={`Coad ${idx+1}`}/>
                    ))}
                </CardContent>
            </Card>
        </div>

        <Card className="shadow-sm">
            <CardHeader className="py-3 flex flex-row justify-between items-center bg-slate-100 border-b">
                <CardTitle className="text-sm font-bold text-slate-700">Tabla de Productos</CardTitle>
                <Button size="sm" onClick={agregarFila} disabled={filas.length >= 9} variant="outline" className="border-dashed border-slate-400">
                    <Plus className="mr-2 h-4 w-4"/> Agregar Lote
                </Button>
            </CardHeader>
            <CardContent className="p-0 overflow-x-auto">
                <table className="w-full text-sm text-left">
                    <thead className="text-xs text-slate-600 uppercase bg-slate-200 border-b">
                        <tr>
                            <th className="px-3 py-2 w-32">Campo</th>
                            <th className="px-2 py-2 w-28 text-center border-l border-slate-300">Has</th>
                            {nombresProductos.map((p, i) => (
                                <th key={i} className="px-1 py-2 text-center text-blue-800 min-w-[80px] border-l border-slate-300">
                                    {p || `P${i+1}`}
                                </th>
                            ))}
                            <th className="w-8"></th>
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                        {filas.map((fila) => (
                            <tr key={fila.id} className="bg-white hover:bg-slate-50">
                                <td className="p-2"><Input value={fila.campo} onChange={(e) => handleFilaChange(fila.id, 'campo', e.target.value)} className="h-8"/></td>
                                <td className="p-2 border-l border-slate-100">
                                    <Input type="number" value={fila.has} onChange={(e) => handleFilaChange(fila.id, 'has', parseFloat(e.target.value))} className="h-8 font-bold text-center w-full min-w-[80px]"/>
                                </td>
                                {fila.dosis.map((d, idx) => (
                                    <td key={idx} className="p-1 border-l border-slate-100">
                                        <Input type="text" inputMode="decimal" value={d === 0 ? '' : d} onChange={(e) => handleFilaChange(fila.id, 'dosis', e.target.value, idx)} className="h-8 text-center text-slate-600" placeholder="-"/>
                                    </td>
                                ))}
                                <td className="p-1 text-center"><button onClick={() => eliminarFila(fila.id)} className="text-slate-300 hover:text-red-500"><Trash2 size={16}/></button></td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </CardContent>
        </Card>

        <Card className="shadow-sm mt-8 border border-yellow-200">
            <CardHeader className="bg-yellow-50/30 py-3">
                <CardTitle className="text-sm font-bold text-yellow-800">Tabla de Coadyuvantes</CardTitle>
            </CardHeader>
            <CardContent className="p-0 overflow-x-auto">
                <table className="w-full text-sm text-left">
                    <thead className="text-xs text-slate-600 uppercase bg-yellow-50 border-b border-yellow-200">
                        <tr>
                            <th className="px-3 py-2 w-32">Lote</th>
                            {nombresCoadyuvantes.map((c, i) => (
                                <th key={i} className="px-2 py-2 text-center border-l border-yellow-200 text-yellow-900 w-1/3">
                                    {c || `Coad ${i+1}`}
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-yellow-100">
                        {filas.map((fila) => (
                            <tr key={fila.id} className="bg-white hover:bg-yellow-50/10">
                                <td className="p-2 font-medium text-slate-700 bg-slate-50 border-r">{fila.campo}</td>
                                {fila.coadyuvantes.map((val, idx) => (
                                    <td key={idx} className="p-2 border-l border-slate-100">
                                        <Input type="text" inputMode="decimal" value={val === 0 ? '' : val} onChange={(e) => handleFilaChange(fila.id, 'coadyuvantes', e.target.value, idx)} className="h-8 text-center border-yellow-100 focus:border-yellow-400" placeholder="-"/>
                                    </td>
                                ))}
                            </tr>
                        ))}
                        {filas.length === 0 && <tr><td colSpan={4} className="p-4 text-center text-slate-400">Sin lotes activos.</td></tr>}
                    </tbody>
                </table>
            </CardContent>
        </Card>

        <Card>
            <CardHeader className="py-2 bg-slate-50"><CardTitle className="text-sm">Observaciones</CardTitle></CardHeader>
            <CardContent className="pt-2">
                <Textarea value={datosGrales.observaciones} onChange={(e) => handleInputChange(e)} rows={3} name="observaciones" className="bg-white"/>
            </CardContent>
        </Card>
      </div>
    </div>
  );
}