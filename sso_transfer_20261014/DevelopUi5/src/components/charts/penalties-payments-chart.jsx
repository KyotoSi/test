
import React, { useState } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, ComposedChart, PieChart, Pie, Cell, AreaChart, Area, Line } from 'recharts';
import { AlertTriangle } from 'lucide-react';

export function PenaltiesPaymentsChart() {
  const [penaltiesPaymentsData] = useState([]);
  const [penaltyTypesData] = useState([]);
  const [supplierPenaltiesData] = useState([]);

  // SAP OData endpoint (для будущей интеграции)

  const EmptyState = ({ message }) => (
    <div className="flex flex-col items-center justify-center h-full bg-gray-50 dark:bg-gray-800/50 rounded-lg min-h-[200px]">
      <AlertTriangle className="w-10 h-10 text-yellow-500 mb-3" />
      <p className="text-gray-500 dark:text-gray-400 text-center">{message}</p>
      <p className="text-gray-400 dark:text-gray-500 text-sm mt-1">Данные появятся после подключения к SAP</p>
    </div>
  );

  // Расчёт метрик с защитой от пустых данных
  const lastData = penaltiesPaymentsData.length > 0 
    ? penaltiesPaymentsData[penaltiesPaymentsData.length - 1]
    : { totalPayments: 0, onTimePayments: 0, averageDelay: 0, penaltyRate: 0 };

  const totalPenalties = penaltiesPaymentsData.reduce((sum, item) => sum + (item.penalties || 0), 0);
  const onTimePercent = lastData.totalPayments > 0 
    ? ((lastData.onTimePayments / lastData.totalPayments) * 100).toFixed(1) 
    : 0;

  return (
    <div className="space-y-6">
      {/* Основной график платежей и пени */}
      <div className="h-80">
        <h4 className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-3">
          Динамика платежей и пени
        </h4>
        {penaltiesPaymentsData.length === 0 ? (
          <EmptyState message="Нет данных о платежах и пени" />
        ) : (
          <ResponsiveContainer width="100%" height="100%">
            <ComposedChart data={penaltiesPaymentsData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" className="opacity-30" />
              <XAxis dataKey="month" tick={{ fontSize: 12 }} />
              <YAxis
                yAxisId="left"
                tick={{ fontSize: 12 }}
                tickFormatter={(value) => `₽${(value / 1000000).toFixed(1)}M`}
              />
              <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 12 }} />
              <Tooltip
                contentStyle={{
                  backgroundColor: 'rgba(255, 255, 255, 0.95)',
                  border: '1px solid #e5e7eb',
                  borderRadius: '8px'
                }}
                formatter={(value, name) => {
                  const labels = {
                    onTimePayments: 'Своевременные платежи',
                    latePayments: 'Просроченные платежи',
                    penalties: 'Пени',
                    penaltyRate: 'Доля пени (%)',
                    savedAmount: 'Сэкономлено'
                  };
                  return [
                    name === 'penaltyRate' ? `${value}%` :
                    typeof value === 'number' && value > 1000 ? `₽${(value / 1000).toFixed(0)}K` : value,
                    labels[name] || name
                  ];
                }}
              />
              <Legend
                formatter={(value) => {
                  const labels = {
                    onTimePayments: 'Своевременные',
                    latePayments: 'Просроченные',
                    penalties: 'Пени',
                    penaltyRate: 'Доля пени',
                    savedAmount: 'Сэкономлено'
                  };
                  return labels[value] || value;
                }}
              />
              <Bar yAxisId="left" dataKey="onTimePayments" fill="#10b981" name="onTimePayments" />
              <Bar yAxisId="left" dataKey="latePayments" fill="#f59e0b" name="latePayments" />
              <Bar yAxisId="left" dataKey="penalties" fill="#ef4444" name="penalties" />
              <Line yAxisId="right" type="monotone" dataKey="penaltyRate" stroke="#8b5cf6" strokeWidth={3} name="penaltyRate" />
              <Line yAxisId="right" type="monotone" dataKey="savedAmount" stroke="#06b6d4" strokeWidth={3} name="savedAmount" />
            </ComposedChart>
          </ResponsiveContainer>
        )}
      </div>

      {/* График среднего времени задержки */}
      <div className="h-64">
        <h4 className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-3">
          Среднее время задержки платежей (дни)
        </h4>
        {penaltiesPaymentsData.length === 0 ? (
          <EmptyState message="Нет данных о задержках" />
        ) : (
          <ResponsiveContainer width="100%" height="100%">
            <AreaChart data={penaltiesPaymentsData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" className="opacity-30" />
              <XAxis dataKey="month" tick={{ fontSize: 12 }} />
              <YAxis tick={{ fontSize: 12 }} />
              <Tooltip
                contentStyle={{
                  backgroundColor: 'rgba(255, 255, 255, 0.95)',
                  border: '1px solid #e5e7eb',
                  borderRadius: '8px'
                }}
                formatter={(value) => [`${value} дн.`, 'Среднее время задержки']}
              />
              <Area
                type="monotone"
                dataKey="averageDelay"
                stroke="#ef4444"
                fill="#ef4444"
                fillOpacity={0.6}
              />
            </AreaChart>
          </ResponsiveContainer>
        )}
      </div>

      {/* Круговые диаграммы */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-gray-50 dark:bg-gray-700/50 p-4 rounded-lg">
          <h4 className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-3">
            Распределение пени по типам нарушений
          </h4>
          <div className="h-64">
            {penaltyTypesData.length === 0 ? (
              <EmptyState message="Нет данных о типах пени" />
            ) : (
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={penaltyTypesData}
                    cx="50%"
                    cy="50%"
                    labelLine={false}
                    label={({ name, value }) => `${name}: ${value}%`}
                    outerRadius={80}
                    dataKey="value"
                  >
                    {penaltyTypesData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.color} />
                    ))}
                  </Pie>
                  <Tooltip
                    formatter={(value, name, props) => [
                      `${value}% (₽${((props.payload.amount || 0) / 1000).toFixed(0)}K)`,
                      'Доля'
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

      {/* Пени по поставщикам */}
      <div className="h-80">
        <h4 className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-3">
          Пени по поставщикам
        </h4>
        {supplierPenaltiesData.length === 0 ? (
          <EmptyState message="Нет данных о пени поставщиков" />
        ) : (
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={supplierPenaltiesData} margin={{ top: 20, right: 30, left: 20, bottom: 60 }}>
              <CartesianGrid strokeDasharray="3 3" className="opacity-30" />
              <XAxis
                dataKey="supplier"
                tick={{ fontSize: 10 }}
                angle={-45}
                textAnchor="end"
                height={80}
              />
              <YAxis
                tick={{ fontSize: 12 }}
                tickFormatter={(value) => `₽${(value / 1000).toFixed(0)}K`}
              />
              <Tooltip
                contentStyle={{
                  backgroundColor: 'rgba(255, 255, 255, 0.95)',
                  border: '1px solid #e5e7eb',
                  borderRadius: '8px'
                }}
                formatter={(value, name) => {
                  const labels = {
                    totalPenalties: 'Общая сумма пени',
                    penaltyCount: 'Количество нарушений',
                    averagePenalty: 'Средняя пеня',
                    improvement: 'Улучшение (%)'
                  };
                  return [
                    name === 'improvement' ? `${value}%` :
                    name === 'penaltyCount' ? value :
                    `₽${(value / 1000).toFixed(0)}K`,
                    labels[name] || name
                  ];
                }}
              />
              <Legend
                formatter={(value) => {
                  const labels = {
                    totalPenalties: 'Общая сумма пени',
                    improvement: 'Улучшение'
                  };
                  return labels[value] || value;
                }}
              />
              <Bar dataKey="totalPenalties" fill="#ef4444" name="totalPenalties" />
            </BarChart>
          </ResponsiveContainer>
        )}
      </div>

      {/* Ключевые метрики */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
        <div className="bg-green-50 dark:bg-green-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-green-600 dark:text-green-400">
            ₽{(lastData.totalPayments / 1000000).toFixed(1)}M
          </div>
          <div className="text-sm text-green-600 dark:text-green-400">Общие платежи</div>
        </div>
        
        <div className="bg-blue-50 dark:bg-blue-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-blue-600 dark:text-blue-400">{onTimePercent}%</div>
          <div className="text-sm text-blue-600 dark:text-blue-400">Своевременно</div>
        </div>
        
        <div className="bg-red-50 dark:bg-red-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-red-600 dark:text-red-400">
            ₽{(totalPenalties / 1000).toFixed(0)}K
          </div>
          <div className="text-sm text-red-600 dark:text-red-400">Всего пени</div>
        </div>
        
        <div className="bg-yellow-50 dark:bg-yellow-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-yellow-600 dark:text-yellow-400">{lastData.averageDelay}</div>
          <div className="text-sm text-yellow-600 dark:text-yellow-400">Дней задержки</div>
        </div>
        
        <div className="bg-purple-50 dark:bg-purple-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-purple-600 dark:text-purple-400">{lastData.penaltyRate}%</div>
          <div className="text-sm text-purple-600 dark:text-purple-400">Доля пени</div>
        </div>
      </div>

      {/* Детальная таблица по поставщикам */}
      <div className="bg-white dark:bg-gray-800 rounded-lg shadow-sm overflow-hidden">
        <div className="px-6 py-4 border-b border-gray-200 dark:border-gray-600">
          <h4 className="text-lg font-medium text-gray-900 dark:text-white">
            Детальная информация по пени поставщиков
          </h4>
        </div>
        {supplierPenaltiesData.length === 0 ? (
          <div className="p-6">
            <EmptyState message="Нет данных о поставщиках" />
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-gray-200 dark:divide-gray-600">
              <thead className="bg-gray-50 dark:bg-gray-700">
                <tr>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Поставщик</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Общая сумма пени</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Количество нарушений</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Средняя пеня</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Улучшение</th>
                </tr>
              </thead>
              <tbody className="bg-white dark:bg-gray-800 divide-y divide-gray-200 dark:divide-gray-600">
                {supplierPenaltiesData.map((supplier, index) => (
                  <tr key={index} className="hover:bg-gray-50 dark:hover:bg-gray-700">
                    <td className="px-6 py-4 text-sm font-medium text-gray-900 dark:text-white">{supplier.supplier}</td>
                    <td className="px-6 py-4 text-sm text-red-600 dark:text-red-400">
                      ₽{((supplier.totalPenalties || 0) / 1000).toFixed(0)}K
                    </td>
                    <td className="px-6 py-4 text-sm text-gray-900 dark:text-white">{supplier.penaltyCount}</td>
                    <td className="px-6 py-4 text-sm text-gray-900 dark:text-white">
                      ₽{((supplier.averagePenalty || 0) / 1000).toFixed(1)}K
                    </td>
                    <td className="px-6 py-4">
                      <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${
                        supplier.improvement > 0 ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'
                      }`}>
                        {supplier.improvement > 0 ? '+' : ''}{supplier.improvement}%
                      </span>
                    </td>
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
