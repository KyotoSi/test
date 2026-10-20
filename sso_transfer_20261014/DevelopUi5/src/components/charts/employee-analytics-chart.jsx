import React, { useState } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, LineChart, Line, PieChart, Pie, Cell } from 'recharts';
import { AlertTriangle } from 'lucide-react';

export function EmployeeAnalyticsChart() {
  const [employeeWorkloadData] = useState([]);
  const [monthlyWorkloadData] = useState([]);
  const [departmentData] = useState([]);

  // SAP OData endpoint (для будущей интеграции)

  const EmptyState = ({ message }) => (
    <div className="flex flex-col items-center justify-center h-full bg-gray-50 dark:bg-gray-800/50 rounded-lg min-h-[200px]">
      <AlertTriangle className="w-10 h-10 text-yellow-500 mb-3" />
      <p className="text-gray-500 dark:text-gray-400 text-center">{message}</p>
      <p className="text-gray-400 dark:text-gray-500 text-sm mt-1">Данные появятся после подключения к SAP</p>
    </div>
  );

  // Расчёт метрик с защитой от деления на ноль
  const totalEmployees = departmentData.reduce((sum, dept) => sum + (dept.employees || 0), 0);
  const avgEfficiency = employeeWorkloadData.length > 0
    ? Math.round(employeeWorkloadData.reduce((sum, emp) => sum + (emp.efficiency || 0), 0) / employeeWorkloadData.length)
    : 0;
  const avgHours = employeeWorkloadData.length > 0
    ? Math.round(employeeWorkloadData.reduce((sum, emp) => sum + (emp.hoursWorked || 0), 0) / employeeWorkloadData.length)
    : 0;
  const avgOvertime = employeeWorkloadData.length > 0
    ? Math.round(employeeWorkloadData.reduce((sum, emp) => sum + (emp.overtime || 0), 0) / employeeWorkloadData.length)
    : 0;
  const avgTasks = employeeWorkloadData.length > 0
    ? Math.round(employeeWorkloadData.reduce((sum, emp) => sum + (emp.tasksCompleted || 0), 0) / employeeWorkloadData.length)
    : 0;

  return (
    <div className="space-y-6">
      {/* Трудозатраты сотрудников */}
      <div className="h-80">
        <h4 className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-3">
          Трудозатраты сотрудников (часы в месяц)
        </h4>
        {employeeWorkloadData.length === 0 ? (
          <EmptyState message="Нет данных о трудозатратах сотрудников" />
        ) : (
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={employeeWorkloadData} margin={{ top: 20, right: 30, left: 20, bottom: 60 }}>
              <CartesianGrid strokeDasharray="3 3" className="opacity-30" />
              <XAxis
                dataKey="employee"
                tick={{ fontSize: 10 }}
                angle={-45}
                textAnchor="end"
                height={80}
              />
              <YAxis tick={{ fontSize: 12 }} />
              <Tooltip
                contentStyle={{
                  backgroundColor: 'rgba(255, 255, 255, 0.95)',
                  border: '1px solid #e5e7eb',
                  borderRadius: '8px'
                }}
                formatter={(value, name) => {
                  const labels = {
                    hoursWorked: 'Отработано часов',
                    overtime: 'Переработка (ч)'
                  };
                  return [value, labels[name] || name];
                }}
              />
              <Legend
                formatter={(value) => {
                  const labels = {
                    hoursWorked: 'Отработано часов',
                    overtime: 'Переработка'
                  };
                  return labels[value] || value;
                }}
              />
              <Bar dataKey="hoursWorked" fill="#3b82f6" name="hoursWorked" />
              <Bar dataKey="overtime" fill="#ef4444" name="overtime" />
            </BarChart>
          </ResponsiveContainer>
        )}
      </div>

      {/* Динамика трудозатрат по месяцам */}
      <div className="h-80">
        <h4 className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-3">
          Динамика трудозатрат команды
        </h4>
        {monthlyWorkloadData.length === 0 ? (
          <EmptyState message="Нет данных о динамике трудозатрат" />
        ) : (
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={monthlyWorkloadData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" className="opacity-30" />
              <XAxis dataKey="month" tick={{ fontSize: 12 }} />
              <YAxis tick={{ fontSize: 12 }} />
              <Tooltip
                contentStyle={{
                  backgroundColor: 'rgba(255, 255, 255, 0.95)',
                  border: '1px solid #e5e7eb',
                  borderRadius: '8px'
                }}
                formatter={(value, name) => {
                  const labels = {
                    totalHours: 'Общие часы',
                    productiveHours: 'Продуктивные часы',
                    overtime: 'Переработка',
                    efficiency: 'Эффективность (%)'
                  };
                  return [
                    name === 'efficiency' ? `${value}%` : `${value} ч`,
                    labels[name] || name
                  ];
                }}
              />
              <Legend
                formatter={(value) => {
                  const labels = {
                    totalHours: 'Общие часы',
                    productiveHours: 'Продуктивные часы',
                    overtime: 'Переработка',
                    efficiency: 'Эффективность'
                  };
                  return labels[value] || value;
                }}
              />
              <Line type="monotone" dataKey="totalHours" stroke="#3b82f6" strokeWidth={3} name="totalHours" />
              <Line type="monotone" dataKey="productiveHours" stroke="#10b981" strokeWidth={3} name="productiveHours" />
              <Line type="monotone" dataKey="overtime" stroke="#ef4444" strokeWidth={3} name="overtime" />
              <Line type="monotone" dataKey="efficiency" stroke="#8b5cf6" strokeWidth={3} name="efficiency" />
            </LineChart>
          </ResponsiveContainer>
        )}
      </div>

      {/* Распределение по отделам */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-gray-50 dark:bg-gray-700/50 p-4 rounded-lg">
          <h4 className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-3">
            Распределение сотрудников по отделам
          </h4>
          <div className="h-64">
            {departmentData.length === 0 ? (
              <EmptyState message="Нет данных по отделам" />
            ) : (
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={departmentData}
                    cx="50%"
                    cy="50%"
                    labelLine={false}
                    label={({ name, employees }) => `${name}: ${employees} чел.`}
                    outerRadius={80}
                    dataKey="value"
                  >
                    {departmentData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.color} />
                    ))}
                  </Pie>
                  <Tooltip
                    formatter={(value, name, props) => [
                      `${props.payload.employees} сотрудников (${value}%)`,
                      'Количество'
                    ]}
                    contentStyle={{
                      backgroundColor: 'rgba(255, 255, 255, 0.95)',
                      border: '1px solid #e5e7eb',
                      borderRadius: '8px'
                    }}
                  />
                </PieChart>
              </ResponsiveContainer>
            )}
          </div>
        </div>
      </div>

      {/* Ключевые метрики сотрудников */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
        <div className="bg-blue-50 dark:bg-blue-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-blue-600 dark:text-blue-400">{totalEmployees}</div>
          <div className="text-sm text-blue-600 dark:text-blue-400">Всего сотрудников</div>
        </div>
        
        <div className="bg-green-50 dark:bg-green-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-green-600 dark:text-green-400">{avgEfficiency}%</div>
          <div className="text-sm text-green-600 dark:text-green-400">Средняя эффективность</div>
        </div>
        
        <div className="bg-yellow-50 dark:bg-yellow-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-yellow-600 dark:text-yellow-400">{avgHours}</div>
          <div className="text-sm text-yellow-600 dark:text-yellow-400">Часов в месяц</div>
        </div>
        
        <div className="bg-red-50 dark:bg-red-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-red-600 dark:text-red-400">{avgOvertime}</div>
          <div className="text-sm text-red-600 dark:text-red-400">Переработка (ч)</div>
        </div>
        
        <div className="bg-purple-50 dark:bg-purple-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-purple-600 dark:text-purple-400">{avgTasks}</div>
          <div className="text-sm text-purple-600 dark:text-purple-400">Задач в месяц</div>
        </div>
      </div>

      {/* Детальная таблица сотрудников */}
      <div className="bg-white dark:bg-gray-800 rounded-lg shadow-sm overflow-hidden">
        <div className="px-6 py-4 border-b border-gray-200 dark:border-gray-600">
          <h4 className="text-lg font-medium text-gray-900 dark:text-white">
            Детальная информация по сотрудникам
          </h4>
        </div>
        {employeeWorkloadData.length === 0 ? (
          <div className="p-6">
            <EmptyState message="Нет данных о сотрудниках" />
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-gray-200 dark:divide-gray-600">
              <thead className="bg-gray-50 dark:bg-gray-700">
                <tr>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Сотрудник</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Отдел</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Часы</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Задачи</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Эффективность</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Переработка</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Проекты</th>
                </tr>
              </thead>
              <tbody className="bg-white dark:bg-gray-800 divide-y divide-gray-200 dark:divide-gray-600">
                {employeeWorkloadData.map((employee, index) => (
                  <tr key={index} className="hover:bg-gray-50 dark:hover:bg-gray-700">
                    <td className="px-6 py-4 text-sm font-medium text-gray-900 dark:text-white">{employee.employee}</td>
                    <td className="px-6 py-4 text-sm text-gray-500 dark:text-gray-300">{employee.department}</td>
                    <td className="px-6 py-4 text-sm text-gray-900 dark:text-white">{employee.hoursWorked}</td>
                    <td className="px-6 py-4 text-sm text-gray-900 dark:text-white">{employee.tasksCompleted}</td>
                    <td className="px-6 py-4">
                      <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${
                        employee.efficiency >= 95 ? 'bg-green-100 text-green-800' :
                        employee.efficiency >= 90 ? 'bg-yellow-100 text-yellow-800' :
                        'bg-red-100 text-red-800'
                      }`}>
                        {employee.efficiency}%
                      </span>
                    </td>
                    <td className="px-6 py-4 text-sm text-gray-900 dark:text-white">{employee.overtime}ч</td>
                    <td className="px-6 py-4 text-sm text-gray-900 dark:text-white">{employee.projects}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
