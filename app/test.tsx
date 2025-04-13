import { useState, useEffect } from 'react';
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
  const [diasSemana, setDiasSemana] = useState({
    Domingo: 0, Lunes: 0, Martes: 0, Miércoles: 0, Jueves: 0, Viernes: 0, Sábado: 0
  });

  // Datos para gráfica por día
  const dataAsistenciaPorDia = [
    { name: 'Domingo', empleados: 1 },
    { name: 'Lunes', empleados: 82 },
    { name: 'Martes', empleados: 96 },
    { name: 'Miércoles', empleados: 104 },
    { name: 'Jueves', empleados: 90 },
    { name: 'Viernes', empleados: 55 },
    { name: 'Sábado', empleados: 0 },
  ];

  // Cargar datos del archivo Excel
  useEffect(() => {
    const cargarDatos = async () => {
      try {
        const response = await fetch('./Records.xls').then(res => res.arrayBuffer());
        console.log("Response:", response);
        
        const workbook = XLSX.read(response, {
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
          return {
            id: emp.id,
            nombre: emp.nombre,
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
        setLoaded(true);
      } catch (error) {
        console.error("Error al cargar datos:", error);
      }
    };
    
    cargarDatos();
  }, []);

  // Filtrar y ordenar empleados
  useEffect(() => {
    if (empleados.length > 0) {
      let filtered = [...empleados];
      
      // Filtrar por búsqueda
      if (searchTerm) {
        filtered = filtered.filter(emp => 
          emp.nombre.toLowerCase().includes(searchTerm.toLowerCase()) ||
          emp.id.includes(searchTerm)
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
        } else if (sortBy === 'nombre') {
          return sortOrder === 'asc'
            ? a.nombre.localeCompare(b.nombre)
            : b.nombre.localeCompare(a.nombre);
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
    return <div className="p-4 text-center">Cargando datos del archivo...</div>;
  }

  return (
    <div className="p-4">
      <h2 className="text-xl font-bold mb-2 text-center">Asistencia de Empleados por Día de la Semana</h2>
      <p className="text-sm mb-4 text-center">Periodo: Marzo 2025 - Total de empleados: {empleados.length}</p>
      
      {/* Gráfica de asistencia por día */}
      <div className="mb-6">
        <h3 className="text-lg font-semibold mb-2">Empleados con asistencia por día</h3>
        <div className="h-64">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={dataAsistenciaPorDia}>
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
            placeholder="Buscar por nombre o ID..."
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