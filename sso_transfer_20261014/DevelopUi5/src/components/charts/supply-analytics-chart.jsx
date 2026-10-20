
import React, { useState } from 'react';
import { ComposedChart, Line, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts';
import { AlertTriangle } from 'lucide-react';

export function SupplyAnalyticsChart() {
  const [supplyAnalyticsData] = useState([]);
  const [supplyCategoryData] = useState([]);
  const [supplierRegionData] = useState([]);
  const [orderDetails] = useState([]);
  const [showDetails, setShowDetails] = useState(false);

  // SAP OData endpoint (для будущей интеграции)

  const EmptyState = ({ message }) => (
    <div className="flex flex-col items-center justify-center h-full bg-gray-50 dark:bg-gray-800/50 rounded-lg min-h-[200px]">
      <AlertTriangle className="w-10 h-10 text-yellow-500 mb-3" />
      <p className="text-gray-500 dark:text-gray-400 text-center">{message}</p>
      <p className="text-gray-400 dark:text-gray-500 text-sm mt-1">Данные появятся после подключения к SAP</p>
    </div>
  );

  // Расчёт метрик с защитой от пустых данных
  const totalSupplies = supplyAnalyticsData.reduce((sum, item) => sum + (item.totalSupplies || 0), 0);
  const urgentSupplies = supplyAnalyticsData.reduce((sum, item) => sum + (item.urgentSupplies || 0), 0);
  const avgDeliveryTime = supplyAnalyticsData.length > 0
    ? (supplyAnalyticsData.reduce((sum, item) => sum + (item.averageDeliveryTime || 0), 0) / supplyAnalyticsData.length).toFixed(1)
    : 0;
  const avgQuality = supplyAnalyticsData.length > 0
    ? (supplyAnalyticsData.reduce((sum, item) => sum + (item.qualityScore || 0), 0) / supplyAnalyticsData.length).toFixed(1)
    : 0;
  const avgCost = supplyAnalyticsData.length > 0
    ? Math.round(supplyAnalyticsData.reduce((sum, item) => sum + (item.costPerSupply || 0), 0) / supplyAnalyticsData.length / 1000)
    : 0;

  return (
    <div className="space-y-6">
      {/* Основной комбинированный график */}
      <div className="h-80">
        {supplyAnalyticsData.length === 0 ? (
          <EmptyState message="Нет данных об аналитике поставок" />
        ) : (
          <ResponsiveContainer width="100%" height="100%">
            <ComposedChart data={supplyAnalyticsData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" className="opacity-30" />
              <XAxis dataKey="month" tick={{ fontSize: 12 }} />
              <YAxis yAxisId="left" tick={{ fontSize: 12 }} />
              <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 12 }} />
              <Tooltip
                contentStyle={{
                  backgroundColor: 'rgba(255, 255, 255, 0.95)',
                  border: '1px solid #e5e7eb',
                  borderRadius: '8px'
                }}
                formatter={(value, name) => {
                  const labels = {
                    totalSupplies: 'Всего поставок',
                    urgentSupplies: 'Срочные поставки',
                    averageDeliveryTime: 'Среднее время доставки (дни)',
                    qualityScore: 'Оценка качества'
                  };
                  return [
                    name === 'averageDeliveryTime' ? `${value} дн.` :
                    name === 'qualityScore' ? `${value}/5` : value,
                    labels[name] || name
                  ];
                }}
              />
              <Legend
                formatter={(value) => {
                  const labels = {
                    totalSupplies: 'Всего поставок',
                    urgentSupplies: 'Срочные поставки',
                    averageDeliveryTime: 'Время доставки',
                    qualityScore: 'Качество'
                  };
                  return labels[value] || value;
                }}
              />
              <Bar yAxisId="left" dataKey="totalSupplies" fill="#3b82f6" name="totalSupplies" />
              <Bar yAxisId="left" dataKey="urgentSupplies" fill="#ef4444" name="urgentSupplies" />
              <Line yAxisId="right" type="monotone" dataKey="averageDeliveryTime" stroke="#f59e0b" strokeWidth={3} name="averageDeliveryTime" />
              <Line yAxisId="right" type="monotone" dataKey="qualityScore" stroke="#10b981" strokeWidth={3} name="qualityScore" />
            </ComposedChart>
          </ResponsiveContainer>
        )}
      </div>

      {/* Круговые диаграммы */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {/* Категории поставок */}
        <div className="bg-gray-50 dark:bg-gray-700/50 p-4 rounded-lg">
          <h4 className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-3">
            Распределение по категориям
          </h4>
          <div className="h-64">
            {supplyCategoryData.length === 0 ? (
              <EmptyState message="Нет данных о категориях" />
            ) : (
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={supplyCategoryData}
                    cx="50%"
                    cy="50%"
                    labelLine={false}
                    label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}
                    outerRadius={80}
                    dataKey="value"
                  >
                    {supplyCategoryData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.color} />
                    ))}
                  </Pie>
                  <Tooltip
                    formatter={(value) => [`${value}%`, 'Доля']}
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

        {/* Регионы поставщиков */}
        <div className="bg-gray-50 dark:bg-gray-700/50 p-4 rounded-lg">
          <h4 className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-3">
            Поставщики по регионам
          </h4>
          <div className="h-64">
            {supplierRegionData.length === 0 ? (
              <EmptyState message="Нет данных о регионах" />
            ) : (
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={supplierRegionData}
                    cx="50%"
                    cy="50%"
                    labelLine={false}
                    label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}
                    outerRadius={80}
                    dataKey="value"
                  >
                    {supplierRegionData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.color} />
                    ))}
                  </Pie>
                  <Tooltip
                    formatter={(value) => [`${value}%`, 'Доля']}
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

      {/* Ключевые метрики */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
        <div className="bg-blue-50 dark:bg-blue-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-blue-600 dark:text-blue-400">{totalSupplies}</div>
          <div className="text-sm text-blue-600 dark:text-blue-400">Всего поставок</div>
        </div>
        
        <div className="bg-red-50 dark:bg-red-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-red-600 dark:text-red-400">{urgentSupplies}</div>
          <div className="text-sm text-red-600 dark:text-red-400">Срочных</div>
        </div>
        
        <div className="bg-yellow-50 dark:bg-yellow-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-yellow-600 dark:text-yellow-400">{avgDeliveryTime}</div>
          <div className="text-sm text-yellow-600 dark:text-yellow-400">Среднее время (дни)</div>
        </div>
        
        <div className="bg-green-50 dark:bg-green-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-green-600 dark:text-green-400">{avgQuality}</div>
          <div className="text-sm text-green-600 dark:text-green-400">Качество (из 5)</div>
        </div>
        
        <div className="bg-purple-50 dark:bg-purple-900/20 p-4 rounded-lg">
          <div className="text-2xl font-bold text-purple-600 dark:text-purple-400">₽{avgCost}K</div>
          <div className="text-sm text-purple-600 dark:text-purple-400">Средняя стоимость</div>
        </div>
      </div>

      {/* Кнопка для показа детальной информации */}
      <div className="mt-6">
        <button
          onClick={() => setShowDetails(!showDetails)}
          disabled={orderDetails.length === 0}
          className="px-4 py-2 bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed text-white rounded-lg transition-colors"
        >
          {showDetails ? 'Скрыть детали' : 'Показать детальную информацию'}
        </button>
      </div>

      {/* Детальная информация о заказах */}
      {showDetails && orderDetails.length > 0 && (
        <div className="mt-6 bg-gray-50 dark:bg-gray-800 rounded-lg p-6">
          <h4 className="text-lg font-semibold text-gray-900 dark:text-white mb-4">
            Детальная информация о поставках
          </h4>
          <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-gray-200 dark:divide-gray-600">
              <thead className="bg-gray-100 dark:bg-gray-700">
                <tr>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Номер заказа</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Поставщик</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">ФИО закупщика</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">ФИО договорника</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Группа</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Статус</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Дата</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Сумма</th>
                </tr>
              </thead>
              <tbody className="bg-white dark:bg-gray-800 divide-y divide-gray-200 dark:divide-gray-600">
                {orderDetails.map((order, index) => (
                  <tr key={index} className="hover:bg-gray-50 dark:hover:bg-gray-700">
                    <td className="px-6 py-4 text-sm font-medium text-gray-900 dark:text-white">{order.orderNumber}</td>
                    <td className="px-6 py-4 text-sm text-gray-500 dark:text-gray-300">{order.supplierName}</td>
                    <td className="px-6 py-4 text-sm text-gray-500 dark:text-gray-300">{order.buyerName}</td>
                    <td className="px-6 py-4 text-sm text-gray-500 dark:text-gray-300">{order.contractSpecialistName}</td>
                    <td className="px-6 py-4 text-sm text-gray-500 dark:text-gray-300">{order.contractSpecialistGroup}</td>
                    <td className="px-6 py-4">
                      <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${
                        order.status === 'Выполнен' ? 'bg-green-100 text-green-800' :
                        order.status === 'Задержан' ? 'bg-yellow-100 text-yellow-800' :
                        order.status === 'В процессе' ? 'bg-blue-100 text-blue-800' :
                        'bg-red-100 text-red-800'
                      }`}>
                        {order.status}
                      </span>
                    </td>
                    <td className="px-6 py-4 text-sm text-gray-500 dark:text-gray-300">
                      {order.deliveryDate ? new Date(order.deliveryDate).toLocaleDateString('ru-RU') : '-'}
                    </td>
                    <td className="px-6 py-4 text-sm text-gray-500 dark:text-gray-300">
                      ₽ {(order.amount || 0).toLocaleString('ru-RU')}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {!showDetails && orderDetails.length === 0 && (
        <div className="mt-4 p-4 bg-yellow-50 dark:bg-yellow-900/20 rounded-lg border border-yellow-200 dark:border-yellow-800">
          <div className="flex items-center gap-3">
            <AlertTriangle className="w-5 h-5 text-yellow-600" />
            <p className="text-yellow-700 dark:text-yellow-400 text-sm">
              Детальная информация появится после подключения к SAP OData
            </p>
          </div>
        </div>
      )}
    </div>
  );
}
