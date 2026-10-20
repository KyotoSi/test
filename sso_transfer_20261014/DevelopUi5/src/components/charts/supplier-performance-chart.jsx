
import React, { useState } from 'react';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Legend,
  Cell
} from 'recharts';
import { AlertTriangle } from 'lucide-react';

export const SupplierPerformanceChart = () => {
  const [data] = useState([]);
  const [sortBy, setSortBy] = useState('rating');

  // SAP OData endpoint (для будущей интеграции)

  const getBarColor = (value, type) => {
    if (type === 'rating') {
      if (value >= 4.5) return '#10B981';
      if (value >= 3.5) return '#F59E0B';
      return '#EF4444';
    }
    if (type === 'onTime') {
      if (value >= 90) return '#10B981';
      if (value >= 70) return '#F59E0B';
      return '#EF4444';
    }
    return '#3B82F6';
  };

  const CustomTooltip = ({ active, payload, label }) => {
    if (active && payload && payload.length) {
      return (
        <div className="bg-white dark:bg-gray-800 p-4 rounded-lg shadow-lg border border-gray-200 dark:border-gray-700">
          <p className="font-semibold text-gray-900 dark:text-white mb-2">{label}</p>
          {payload.map((item, index) => (
            <p key={index} className="text-sm text-gray-600 dark:text-gray-300">
              {item.name}: {item.value}
              {item.dataKey === 'onTimePercentage' ? '%' : ''}
            </p>
          ))}
        </div>
      );
    }
    return null;
  };

  return (
    <div className="w-full">
      {/* Фильтры сортировки */}
      <div className="flex gap-2 mb-4">
        <button
          onClick={() => setSortBy('rating')}
          className={`px-3 py-1 text-sm rounded-lg transition-colors ${
            sortBy === 'rating'
              ? 'bg-blue-600 text-white'
              : 'bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300 hover:bg-gray-200'
          }`}
        >
          По рейтингу
        </button>
        <button
          onClick={() => setSortBy('onTime')}
          className={`px-3 py-1 text-sm rounded-lg transition-colors ${
            sortBy === 'onTime'
              ? 'bg-blue-600 text-white'
              : 'bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300 hover:bg-gray-200'
          }`}
        >
          По своевременности
        </button>
        <button
          onClick={() => setSortBy('orders')}
          className={`px-3 py-1 text-sm rounded-lg transition-colors ${
            sortBy === 'orders'
              ? 'bg-blue-600 text-white'
              : 'bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300 hover:bg-gray-200'
          }`}
        >
          По заказам
        </button>
      </div>

      {/* График или заглушка */}
      {data.length === 0 ? (
        <div className="flex flex-col items-center justify-center h-[500px] bg-gray-50 dark:bg-gray-800/50 rounded-lg">
          <AlertTriangle className="w-12 h-12 text-yellow-500 mb-4" />
          <p className="text-gray-500 dark:text-gray-400 text-center text-lg">
            Нет данных о поставщиках
          </p>
          <p className="text-gray-400 dark:text-gray-500 text-sm mt-2">
            Данные появятся после подключения к SAP OData
          </p>
        </div>
      ) : (
        <ResponsiveContainer width="100%" height={500}>
          <BarChart
            data={data}
            layout="vertical"
            margin={{ top: 5, right: 30, left: 100, bottom: 5 }}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#E5E7EB" />
            <XAxis type="number" stroke="#9CA3AF" />
            <YAxis dataKey="name" type="category" stroke="#9CA3AF" width={100} />
            <Tooltip content={<CustomTooltip />} />
            <Legend />
            <Bar dataKey="rating" name="Рейтинг" fill="#3B82F6">
              {data.map((entry, index) => (
                <Cell key={`cell-${index}`} fill={getBarColor(entry.rating, 'rating')} />
              ))}
            </Bar>
            <Bar dataKey="onTimePercentage" name="Своевременность %" fill="#10B981">
              {data.map((entry, index) => (
                <Cell key={`cell-${index}`} fill={getBarColor(entry.onTimePercentage, 'onTime')} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      )}

      {/* Легенда статусов */}
      <div className="mt-4 flex flex-wrap gap-4 justify-center">
        <div className="flex items-center gap-2">
          <div className="w-4 h-4 rounded bg-green-500"></div>
          <span className="text-sm text-gray-600 dark:text-gray-400">Отлично (≥90%)</span>
        </div>
        <div className="flex items-center gap-2">
          <div className="w-4 h-4 rounded bg-yellow-500"></div>
          <span className="text-sm text-gray-600 dark:text-gray-400">Хорошо (70-90%)</span>
        </div>
        <div className="flex items-center gap-2">
          <div className="w-4 h-4 rounded bg-red-500"></div>
          <span className="text-sm text-gray-600 dark:text-gray-400">Требует внимания (&lt;70%)</span>
        </div>
      </div>
    </div>
  );
};

export default SupplierPerformanceChart;
