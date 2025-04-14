import { useState, useEffect, useRef } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import * as XLSX from 'xlsx';

const AsistenciaEmpleados = () => {
  const [empleados, setEmpleados] = useState([]);
  const [filteredEmpleados, setFilteredEmpleados] = useState([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [sortBy, setSortBy] = useState('id');
  const [sortOrder, setSortOrder] = useState('asc');
  const [showOnlyActive, setShowOnlyActive] = useState(true);
  const [loaded, setLoaded] = useState(false);
  const [loading, setLoading] = useState(false);
  const [diasSemana, setDiasSemana] = useState({
    Domingo: 0, Lunes: 0, Martes: 0, Miércoles: 0, Jueves: 0, Viernes: 0, Sábado: 0
  });
  const [dataBarChart, setDataBarChart] = useState([
    { name: 'Domingo', empleados: 0 },
    { name: 'Lunes', empleados: 0 },
    { name: 'Martes', empleados: 0 },
    { name: 'Miércoles', empleados: 0 },
    { name: 'Jueves', empleados: 0 },
    { name: 'Viernes', empleados: 0 },
    { name: 'Sábado', empleados: 0 },
  ]);
  const [error, setError] = useState(null);
  const [recordsUploaded, setRecordsUploaded] = useState(false);
  const [libroUploaded, setLibroUploaded] = useState(false);
  
  // Refs para los inputs de archivo
  const recordsInputRef = useRef(null);
  const libroInputRef = useRef(null);

  // Procesar archivos cargados
  const processFiles = async () => {
    try {
      setLoading(true);
      setError(null);
      
      // Verificar si ambos archivos han sido cargados
      if (!recordsUploaded || !libroUploaded) {
        setError("Por favor carga ambos archivos para procesar los datos.");
        setLoading(false);
        return;
      }
      
      // 1. Procesar Libro15
      const infoAdicionalEmpleados = {};
      
      if (libroInputRef.current.files[0]) {
        const libroFile = await libroInputRef.current.files[0].arrayBuffer();
        
        const workbook = XLSX.read(libroFile, {
          cellStyles: true,
          cellFormulas: true,
          cellDates: true,
          cellNF: true,
          sheetStubs: true
        });
        
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1});
        
        // Buscar la fila de encabezados
        let headerRow = -1;
        for (let i = 0; i < 10 && headerRow === -1; i++) {
          if (data[i] && 
             (data[i].includes("No. Empleado") || 
              data[i].some(cell => cell === "Área" || cell === "Area"))) {
            headerRow = i;
          }
        }
        
        // Obtener índices de columnas
        let indiceID = -1;
        let indiceArea = -1;
        let indicePuesto = -1;
        let indiceRangoTiempo = -1;
        let indiceRangoKM = -1;
        let indiceDiasOficina = -1;
        
        if (headerRow !== -1) {
          const headers = data[headerRow];
          for (let i = 0; i < headers.length; i++) {
            const header = headers[i];
            if (header === "No. Empleado" || header === "ID") indiceID = i;
            if (header === "Área" || header === "Area") indiceArea = i;
            if (header === "Puesto") indicePuesto = i;
            if (header === "Rango Tiempo") indiceRangoTiempo = i;
            if (header === "Rango KM") indiceRangoKM = i;
            if (header === "Días en Oficina") indiceDiasOficina = i;
          }
        }
        
        // Extraer datos si encontramos las columnas
        if (headerRow !== -1 && indiceID !== -1) {
          for (let i = headerRow + 1; i < data.length; i++) {
            const row = data[i];
            if (row && row[indiceID]) {
              let id = row[indiceID];
              // Convertir a string si es número
              if (typeof id === 'number') {
                id = id.toString();
              }
              
              // Si el ID tiene el formato "123-Nombre", extraer solo el número
              if (typeof id === 'string' && id.includes('-')) {
                id = id.split('-')[0];
              }
              
              infoAdicionalEmpleados[id] = {
                area: row[indiceArea] || "",
                puesto: row[indicePuesto] || "",
                rangoTiempo: row[indiceRangoTiempo] || "",
                rangoKM: row[indiceRangoKM] || "",
                diasOficina: indiceDiasOficina !== -1 ? row[indiceDiasOficina] || "" : ""
              };
            }
          }
        }
        
        console.log(`Información adicional cargada para ${Object.keys(infoAdicionalEmpleados).length} empleados`);
      }
      
      // 2. Procesar Records
      if (recordsInputRef.current.files[0]) {
        const recordsFile = await recordsInputRef.current.files[0].arrayBuffer();
        
        const workbook = XLSX.read(recordsFile, {
          cellStyles: true,
          cellFormulas: true,
          cellDates: true,
          cellNF: true,
          sheetStubs: true
        });
        
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, raw: false, defval: ''});
        
        // Función para convertir fecha a día de la semana
        function getDayOfWeek(dateStr) {
          let parts;
          
          if (dateStr.match(/^\d{1,2}\/\d{1,2}\/\d{2,4}$/)) {
            parts = dateStr.split('/');
          } else if (dateStr.match(/^\d{1,2}-\d{1,2}-\d{2,4}$/)) {
            parts = dateStr.split('-');
          } else {
            return null;
          }
          
          let year = parseInt(parts[2]);
          if (year < 100) {
            year += 2000;
          }
          
          const date = new Date(year, parseInt(parts[1]) - 1, parseInt(parts[0]));
          return date.getDay();
        }
        
        const employeeAttendance = {};
        let currentEmployee = null;
        let employeeName = "";
        let employeeId = "";
        
        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          
          if (row[0] && row[0].match(/^\d+-/)) {
            const match = row[0].match(/^(\d+)-(.+)$/);
            if (match) {
              employeeId = match[1];
              employeeName = match[2].trim();
              currentEmployee = employeeName;
              
              if (!employeeAttendance[currentEmployee]) {
                employeeAttendance[currentEmployee] = {
                  id: employeeId,
                  nombre: currentEmployee,
                  0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0
                };
              }
            }
            continue;
          }
          
          if (row[0] && (row[0].match(/^\d{1,2}\/\d{1,2}\/\d{2,4}$/) || row[0].match(/^\d{1,2}-\d{1,2}-\d{2,4}$/)) && currentEmployee) {
            if (row[1] === "CX7") {
              const dayOfWeek = getDayOfWeek(row[0]);
              if (dayOfWeek !== null && employeeAttendance[currentEmployee]) {
                employeeAttendance[currentEmployee][dayOfWeek]++;
              }
            }
          }
        }
        
        // Convertir a array para ordenar
        const employeeAttendanceArray = Object.values(employeeAttendance);
        
        // Preparar formato para la tabla
        const formattedData = employeeAttendanceArray.map(emp => {
          // Buscar información adicional del empleado
          const infoAdicional = infoAdicionalEmpleados[emp.id] || {};
          
          return {
            id: emp.id,
            nombre: emp.nombre,
            area: infoAdicional.area || "No disponible",
            puesto: infoAdicional.puesto || "No disponible",
            rangoTiempo: infoAdicional.rangoTiempo || "No disponible",
            rangoKM: infoAdicional.rangoKM || "No disponible",
            diasOficina: infoAdicional.diasOficina || "No disponible",
            Domingo: emp[0],
            Lunes: emp[1],
            Martes: emp[2],
            Miércoles: emp[3],
            Jueves: emp[4],
            Viernes: emp[5],
            Sábado: emp[6],
            total: emp[0] + emp[1] + emp[2] + emp[3] + emp[4] + emp[5] + emp[6]
          };
        });
        
        // Ordenar por ID numérico inicialmente
        formattedData.sort((a, b) => parseInt(a.id) - parseInt(b.id));
        
        setEmpleados(formattedData);
        
        // Actualizar datos para el gráfico
        const diasCount = {
          Domingo: formattedData.filter(e => e.Domingo > 0).length,
          Lunes: formattedData.filter(e => e.Lunes > 0).length,
          Martes: formattedData.filter(e => e.Martes > 0).length,
          Miércoles: formattedData.filter(e => e.Miércoles > 0).length,
          Jueves: formattedData.filter(e => e.Jueves > 0).length,
          Viernes: formattedData.filter(e => e.Viernes > 0).length,
          Sábado: formattedData.filter(e => e.Sábado > 0).length
        };
        
        const chartData = [
          { name: 'Domingo', empleados: diasCount.Domingo },
          { name: 'Lunes', empleados: diasCount.Lunes },
          { name: 'Martes', empleados: diasCount.Martes },
          { name: 'Miércoles', empleados: diasCount.Miércoles },
          { name: 'Jueves', empleados: diasCount.Jueves },
          { name: 'Viernes', empleados: diasCount.Viernes },
          { name: 'Sábado', empleados: diasCount.Sábado },
        ];
        
        setDataBarChart(chartData);
        setLoaded(true);
      }
    } catch (error) {
      console.error("Error al procesar archivos:", error);
      setError("Error al procesar los archivos. Verifica que los archivos tengan el formato correcto.");
    } finally {
      setLoading(false);
    }
  };

  // Manejar carga de archivos
  const handleRecordsUpload = (e) => {
    if (e.target.files && e.target.files[0]) {
      setRecordsUploaded(true);
    } else {
      setRecordsUploaded(false);
    }
  };

  const handleLibroUpload = (e) => {
    if (e.target.files && e.target.files[0]) {
      setLibroUploaded(true);
    } else {
      setLibroUploaded(false);
    }
  };

  // Filtrar y ordenar empleados
  useEffect(() => {
    if (empleados.length > 0) {
      let filtered = [...empleados];
      
      // Filtrar por búsqueda
      if (searchTerm) {
        filtered = filtered.filter(emp => 
          emp.nombre.toLowerCase().includes(searchTerm.toLowerCase()) ||
          emp.id.includes(searchTerm) ||
          emp.area.toLowerCase().includes(searchTerm.toLowerCase()) ||
          emp.puesto.toLowerCase().includes(searchTerm.toLowerCase()) ||
          emp.diasOficina.toLowerCase().includes(searchTerm.toLowerCase())
        );
      }
      
      // Filtrar solo activos
      if (showOnlyActive) {
        filtered = filtered.filter(emp => emp.total > 0);
      }
      
      // Ordenar
      filtered.sort((a, b) => {
        if (sortBy === 'id') {
          return sortOrder === 'asc' 
            ? parseInt(a.id) - parseInt(b.id)
            : parseInt(b.id) - parseInt(a.id);
        } else if (sortBy === 'nombre' || sortBy === 'area' || sortBy === 'puesto' ||
                  sortBy === 'rangoTiempo' || sortBy === 'rangoKM' || sortBy === 'diasOficina') {
          return sortOrder === 'asc'
            ? a[sortBy].localeCompare(b[sortBy])
            : b[sortBy].localeCompare(a[sortBy]);
        } else if (sortBy === 'total') {
          return sortOrder === 'asc'
            ? a.total - b.total
            : b.total - a.total;
        } else {
          // Para días de la semana
          return sortOrder === 'asc'
            ? a[sortBy] - b[sortBy]
            : b[sortBy] - a[sortBy];
        }
      });
      
      setFilteredEmpleados(filtered);

      // Calcular totales por día
      const diasCount = {
        Domingo: filtered.filter(e => e.Domingo > 0).length,
        Lunes: filtered.filter(e => e.Lunes > 0).length,
        Martes: filtered.filter(e => e.Martes > 0).length,
        Miércoles: filtered.filter(e => e.Miércoles > 0).length,
        Jueves: filtered.filter(e => e.Jueves > 0).length,
        Viernes: filtered.filter(e => e.Viernes > 0).length,
        Sábado: filtered.filter(e => e.Sábado > 0).length
      };
      
      setDiasSemana(diasCount);
    }
  }, [empleados, searchTerm, sortBy, sortOrder, showOnlyActive]);

  // Manejar click en encabezado para ordenar
  const handleSort = (column) => {
    if (sortBy === column) {
      setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
    } else {
      setSortBy(column);
      setSortOrder('desc'); // Por defecto ordenar descendente
    }
  };

  // Renderizar icono de ordenamiento
  const renderSortIcon = (column) => {
    if (sortBy !== column) return null;
    return sortOrder === 'asc' ? ' ▲' : ' ▼';
  };

  if (!loaded) {
    return (
      <div className="p-8 max-w-4xl mx-auto">
        <h2 className="text-2xl font-bold mb-6 text-center">Asistencia de Empleados por Día de la Semana</h2>
        
        {error && (
          <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded mb-4">
            <p>{error}</p>
          </div>
        )}
        
        <div className="bg-gray-50 p-6 rounded-lg shadow-md mb-6">
          <h3 className="text-lg font-semibold mb-4">Cargar archivos</h3>
          <p className="mb-4 text-gray-600">
            Para generar el reporte de asistencia, necesitas cargar los siguientes archivos:
          </p>
          
          <div className="space-y-4">
            <div>
              <label className="block mb-2 font-medium">
                1. Archivo de asistencia (Records_AllDepts):
              </label>
              <div className="flex items-center">
                <input 
                  type="file" 
                  accept=".xls,.xlsx"
                  ref={recordsInputRef}
                  onChange={handleRecordsUpload}
                  className="block w-full text-sm text-gray-500
                    file:mr-4 file:py-2 file:px-4
                    file:rounded-md file:border-0
                    file:text-sm file:font-semibold
                    file:bg-blue-50 file:text-blue-700
                    hover:file:bg-blue-100"
                />
                {recordsUploaded && (
                  <span className="ml-2 text-green-600">✓</span>
                )}
              </div>
              <p className="mt-1 text-xs text-gray-500">
                Archivo con el registro de asistencias diarias.
              </p>
            </div>
            
            <div>
              <label className="block mb-2 font-medium">
                2. Archivo de información adicional (Libro15):
              </label>
              <div className="flex items-center">
                <input 
                  type="file" 
                  accept=".xls,.xlsx"
                  ref={libroInputRef}
                  onChange={handleLibroUpload}
                  className="block w-full text-sm text-gray-500
                    file:mr-4 file:py-2 file:px-4
                    file:rounded-md file:border-0
                    file:text-sm file:font-semibold
                    file:bg-blue-50 file:text-blue-700
                    hover:file:bg-blue-100"
                />
                {libroUploaded && (
                  <span className="ml-2 text-green-600">✓</span>
                )}
              </div>
              <p className="mt-1 text-xs text-gray-500">
                Archivo con información de áreas, puestos, rangos de tiempo y KM.
              </p>
            </div>
          </div>
          
          <div className="mt-6">
            <button
              onClick={processFiles}
              disabled={loading || !recordsUploaded || !libroUploaded}
              className={`w-full py-2 px-4 rounded-md font-medium text-white 
                ${(!recordsUploaded || !libroUploaded) 
                  ? 'bg-gray-400 cursor-not-allowed' 
                  : 'bg-blue-600 hover:bg-blue-700'}`}
            >
              {loading ? 'Procesando...' : 'Generar reporte'}
            </button>
          </div>
        </div>
        
        <div className="text-sm text-gray-600">
          <h4 className="font-semibold mb-2">Instrucciones:</h4>
          <ol className="list-decimal pl-5 space-y-1">
            <li>Carga el archivo de asistencia (Records_AllDepts) con los registros diarios.</li>
            <li>Carga el archivo de información complementaria (Libro15) con datos de áreas y puestos.</li>
            <li>Haz clic en "Generar reporte" para procesar la información.</li>
          </ol>
        </div>
      </div>
    );
  }

  return (
    <div className="p-4">
      <h2 className="text-xl font-bold mb-2 text-center">Asistencia de Empleados por Día de la Semana</h2>
      <p className="text-sm mb-4 text-center">Total de empleados: {empleados.length}</p>
      
      {/* Botón para nueva carga */}
      <div className="mb-4 flex justify-center">
        <button 
          onClick={() => {
            setLoaded(false);
            setEmpleados([]);
            setFilteredEmpleados([]);
            setRecordsUploaded(false);
            setLibroUploaded(false);
            if (recordsInputRef.current) recordsInputRef.current.value = '';
            if (libroInputRef.current) libroInputRef.current.value = '';
          }}
          className="px-4 py-2 bg-gray-200 hover:bg-gray-300 rounded"
        >
          Cargar nuevos archivos
        </button>
      </div>
      
      {/* Gráfica de asistencia por día */}
      <div className="mb-6">
        <h3 className="text-lg font-semibold mb-2">Empleados con asistencia por día</h3>
        <div className="h-64">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={dataBarChart}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis />
              <Tooltip formatter={(value) => [`${value} empleados`, 'Empleados']} />
              <Legend />
              <Bar dataKey="empleados" name="Empleados" fill="#3b82f6" />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div> t
      
      <p className="ext-center">Periodo: Marzo 2025 - Total de empleados: {empleados.length}</p>
      
      {/* Gráfica de asistencia por día */}
      <div className="mb-6">
        <h3 className="text-lg font-semibold mb-2">Empleados con asistencia por día</h3>
        <div className="h-64">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={dataBarChart}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis />
              <Tooltip formatter={(value) => [`${value} empleados`, 'Empleados']} />
              <Legend />
              <Bar dataKey="empleados" name="Empleados" fill="#3b82f6" />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
      
      {/* Controles de filtrado */}
      <div className="flex flex-col md:flex-row gap-4 mb-4">
        <div className="flex-1">
          <input
            type="text"
            placeholder="Buscar por nombre, ID, área o puesto..."
            className="w-full p-2 border rounded"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
        <div className="flex items-center">
          <label className="flex items-center cursor-pointer">
            <input
              type="checkbox"
              checked={showOnlyActive}
              onChange={() => setShowOnlyActive(!showOnlyActive)}
              className="mr-2"
            />
            <span>Mostrar solo con asistencia</span>
          </label>
        </div>
      </div>
      
      {/* Resumen */}
      <div className="bg-gray-100 p-3 rounded-lg mb-4">
        <h3 className="font-semibold mb-2">Resumen:</h3>
        <ul className="grid grid-cols-1 md:grid-cols-2 gap-2">
          <li>Total de empleados: <span className="font-semibold">{empleados.length}</span></li>
          <li>Con al menos una asistencia: <span className="font-semibold">{empleados.filter(e => e.total > 0).length}</span></li>
          <li>Con asistencia completa (20+ días): <span className="font-semibold">{empleados.filter(e => e.total >= 20).length}</span></li>
          <li>Mostrando: <span className="font-semibold">{filteredEmpleados.length} empleados</span></li>
        </ul>
      </div>
      
      {/* Tabla de empleados */}
      <div className="overflow-x-auto">
        <table className="min-w-full bg-white border">
          <thead className="bg-gray-100">
            <tr>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('id')}>
                ID {renderSortIcon('id')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('nombre')}>
                Nombre {renderSortIcon('nombre')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('area')}>
                Área {renderSortIcon('area')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('puesto')}>
                Puesto {renderSortIcon('puesto')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('rangoTiempo')}>
                Rango Tiempo {renderSortIcon('rangoTiempo')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('rangoKM')}>
                Rango KM {renderSortIcon('rangoKM')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('diasOficina')}>
                Días en Oficina {renderSortIcon('diasOficina')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('Domingo')}>
                Dom {renderSortIcon('Domingo')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('Lunes')}>
                Lun {renderSortIcon('Lunes')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('Martes')}>
                Mar {renderSortIcon('Martes')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('Miércoles')}>
                Mié {renderSortIcon('Miércoles')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('Jueves')}>
                Jue {renderSortIcon('Jueves')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('Viernes')}>
                Vie {renderSortIcon('Viernes')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('Sábado')}>
                Sáb {renderSortIcon('Sábado')}
              </th>
              <th className="px-4 py-2 cursor-pointer" onClick={() => handleSort('total')}>
                Total {renderSortIcon('total')}
              </th>
            </tr>
          </thead>
          <tbody>
            {filteredEmpleados.map((emp, index) => (
              <tr key={emp.id} className={index % 2 === 0 ? 'bg-gray-50' : ''}>
                <td className="border px-4 py-2">{emp.id}</td>
                <td className="border px-4 py-2">{emp.nombre}</td>
                <td className="border px-4 py-2">{emp.area}</td>
                <td className="border px-4 py-2">{emp.puesto}</td>
                <td className="border px-4 py-2">{emp.rangoTiempo}</td>
                <td className="border px-4 py-2">{emp.rangoKM}</td>
                <td className="border px-4 py-2">{emp.diasOficina}</td>
                <td className={`border px-4 py-2 text-center ${emp.Domingo > 0 ? 'bg-green-100' : ''}`}>{emp.Domingo}</td>
                <td className={`border px-4 py-2 text-center ${emp.Lunes > 0 ? 'bg-green-100' : ''}`}>{emp.Lunes}</td>
                <td className={`border px-4 py-2 text-center ${emp.Martes > 0 ? 'bg-green-100' : ''}`}>{emp.Martes}</td>
                <td className={`border px-4 py-2 text-center ${emp.Miércoles > 0 ? 'bg-green-100' : ''}`}>{emp.Miércoles}</td>
                <td className={`border px-4 py-2 text-center ${emp.Jueves > 0 ? 'bg-green-100' : ''}`}>{emp.Jueves}</td>
                <td className={`border px-4 py-2 text-center ${emp.Viernes > 0 ? 'bg-green-100' : ''}`}>{emp.Viernes}</td>
                <td className={`border px-4 py-2 text-center ${emp.Sábado > 0 ? 'bg-green-100' : ''}`}>{emp.Sábado}</td>
                <td className={`border px-4 py-2 text-center font-bold ${emp.total >= 20 ? 'bg-blue-200' : emp.total > 0 ? 'bg-blue-100' : ''}`}>
                  {emp.total}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      
      <div className="mt-4 text-sm text-gray-600">
        <p>* Haz clic en los encabezados de columna para ordenar la tabla.</p>
        <p>* El total muestra la suma de días con asistencia durante el mes.</p>
        <p>* Las celdas en verde indican días con asistencia.</p>
        <p>* Los empleados con asistencia completa (20+ días) se muestran con un total resaltado en azul.</p>
      </div>
    </div>
  );
};

export default AsistenciaEmpleados;