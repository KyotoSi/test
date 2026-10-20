
import React, { useState } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from 'recharts';
import { AlertTriangle } from 'lucide-react';

export function ProcessingSpeedChart() {
  const [data] = useState(null);
  const [isLoading] = useState(false);

  // SAP OData endpoint (для будущей интеграции)

  if (isLoading) {
    return (
      <div className="flex items-center justify-center h-full">
        <div className="text-gray-500">Загрузка данных...</div>
      </div>
    );
  }

  if (!data) {
    return (
      <div className="h-full flex flex-col">
        {/* Key metrics - пустые */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
          <div className="bg-gradient-to-r from-blue-50 to-cyan-50 dark:from-blue-900/20 dark:to-cyan-900/20 p-4 rounded-lg">
            <p className="text-sm text-blue-600 dark:text-blue-400 font-medium">Среднее время обработки заказа</p>
            <p className="text-2xl font-bold text-blue-900 dark:text-blue-100">0 дня</p>
          </div>
          <div className="bg-gradient-to-r from-green-50 to-emerald-50 dark:from-green-900/20 dark:to-emerald-900/20 p-4 rounded-lg">
            <p className="text-sm text-green-600 dark:text-green-400 font-medium">Процент своевременных поставок</p>
            <p className="text-2xl font-bold text-green-900 dark:text-green-100">0%</p>
          </div>
        </div>

        {/* Заглушка */}
        <div className="flex-1 flex flex-col items-center justify-center bg-gray-50 dark:bg-gray-800/50 rounded-lg p-8">
          <AlertTriangle className="w-12 h-12 text-yellow-500 mb-4" />
          <p className="text-gray-500 dark:text-gray-400 text-center text-lg">Нет данных о скорости обработки</p>
          <p className="text-gray-400 dark:text-gray-500 text-sm mt-2 text-center">
            Данные появятся после подключения к SAP OData
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="h-full flex flex-col">
      {/* Key metrics */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
        <div className="bg-gradient-to-r from-blue-50 to-cyan-50 dark:from-blue-900/20 dark:to-cyan-900/20 p-4 rounded-lg">
          <p className="text-sm text-blue-600 dark:text-blue-400 font-medium">{data.avg_processing_time.description}</p>
          <p className="text-2xl font-bold text-blue-900 dark:text-blue-100">{data.avg_processing_time.value} {data.avg_processing_time.unit}</p>
        </div>
        <div className="bg-gradient-to-r from-green-50 to-emerald-50 dark:from-green-900/20 dark:to-emerald-900/20 p-4 rounded-lg">
          <p className="text-sm text-green-600 dark:text-green-400 font-medium">{data.on_time_delivery_rate.description}</p>
          <p className="text-2xl font-bold text-green-900 dark:text-green-100">{data.on_time_delivery_rate.value}{data.on_time_delivery_rate.unit}</p>
        </div>
      </div>

      {/* Charts */}
      <div className="flex-1 min-h-0 flex flex-col space-y-4">
        <div className="flex-1">
          <h4 className="text-md font-semibold text-gray-900 dark:text-white mb-2">Среднее время обработки по категориям</h4>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart
              data={data.processing_time_by_category}
              margin={{ top: 5, right: 30, left: 20, bottom: 60 }}
            >
              <CartesianGrid strokeDasharray="3 3" strokeOpacity={0.3} />
              <XAxis
                dataKey="category"
                stroke="#6B7280"
                tick={{ fill: '#6B7280', fontSize: 10 }}
                angle={-45}
                textAnchor="end"
                height={60}
                interval={0}
              />
              <YAxis stroke="#6B7280" tick={{ fill: '#6B7280' }} tickFormatter={(value) => `${value} дня`} />
              <Tooltip
                contentStyle={{
                  backgroundColor: '#ffffff',
                  borderColor: '#e5e7eb',
                  borderRadius: '0.5rem',
                  boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.1)'
                }}
                labelStyle={{ color: '#374151' }}
                itemStyle={{ color: '#6b7280' }}
              />
              <Bar dataKey="avg_time" fill="#3B82F6" name="Среднее время (дни)" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div className="flex-1">
          <h4 className="text-md font-semibold text-gray-900 dark:text-white mb-2">Процент своевременных поставок по поставщикам</h4>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart
              data={data.on_time_delivery_by_supplier}
              margin={{ top: 5, right: 30, left: 20, bottom: 60 }}
            >
              <CartesianGrid strokeDasharray="3 3" strokeOpacity={0.3} />
              <XAxis
                dataKey="supplier"
                stroke="#6B7280"
                tick={{ fill: '#6B7280', fontSize: 10 }}
                angle={-45}
                textAnchor="end"
                height={60}
                interval={0}
              />
              <YAxis stroke="#6B7280" tick={{ fill: '#6B7280' }} tickFormatter={(value) => `${value}%`} domain={[0, 100]} />
              <Tooltip
                contentStyle={{
                  backgroundColor: '#ffffff',
                  borderColor: '#e5e7eb',
                  borderRadius: '0.5rem',
                  boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.1)'
                }}
                labelStyle={{ color: '#374151' }}
                itemStyle={{ color: '#6b7280' }}
              />
              <Bar dataKey="on_time_rate" fill="#10B981" name="Процент своевременных поставок (%)" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}
