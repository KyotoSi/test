
import React, { useState } from 'react';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip } from 'recharts';
import { FileText, Zap, TrendingUp, Archive, AlertTriangle } from 'lucide-react';

const RADIAN = Math.PI / 180;

const renderCustomizedLabel = ({ cx, cy, midAngle, innerRadius, outerRadius, percentage }) => {
  const radius = innerRadius + (outerRadius - innerRadius) * 0.5;
  const x = cx + radius * Math.cos(-midAngle * RADIAN);
  const y = cy + radius * Math.sin(-midAngle * RADIAN);

  return (
    <text
      x={x}
      y={y}
      fill="white"
      textAnchor={x > cx ? 'start' : 'end'}
      dominantBaseline="central"
      fontSize="12"
      fontWeight="bold"
      style={{ textShadow: '1px 1px 2px rgba(0,0,0,0.8)' }}
    >
      {`${(percentage || 0).toFixed(1)}%`}
    </text>
  );
};

export function DocumentTypesChart() {
  const [data] = useState([]);
  const [totalDocuments] = useState(0);
  const [edoEfficiency] = useState(0);
  const [isLoading] = useState(false);
  const [hoveredSegment, setHoveredSegment] = useState(null);
  const [showDetails, setShowDetails] = useState(false);
  const [documentDetails] = useState([]);

  // SAP OData endpoint (для будущей интеграции)

  const CustomTooltip = ({ active, payload }) => {
    if (active && payload && payload.length) {
      const data = payload[0].payload;
      return (
        <div className="bg-white dark:bg-gray-800 p-4 border border-gray-200 dark:border-gray-600 rounded-lg shadow-xl">
          <div className="flex items-center gap-2 mb-3">
            {data.type === 'ЭДО' ? (
              <Zap className="w-5 h-5 text-green-500" />
            ) : (
              <FileText className="w-5 h-5 text-amber-500" />
            )}
            <span className="font-semibold text-gray-900 dark:text-white text-lg">
              {data.name}
            </span>
          </div>
          <p className="text-sm text-gray-600 dark:text-gray-300 mb-3">
            {data.description}
          </p>
          <div className="space-y-2 text-sm">
            <div className="flex justify-between">
              <span className="text-gray-600 dark:text-gray-300">Количество:</span>
              <span className="font-semibold text-gray-900 dark:text-white">{data.count}</span>
            </div>
            <div className="flex justify-between">
              <span className="text-gray-600 dark:text-gray-300">Доля:</span>
              <span className="font-semibold text-gray-900 dark:text-white">{data.percentage}%</span>
            </div>
          </div>
        </div>
      );
    }
    return null;
  };

  const CustomLegend = () => {
    return (
      <div className="mt-6 space-y-3">
        {data.map((entry, index) => {
          const isHovered = hoveredSegment === entry.type;
          return (
            <div
              key={`legend-${index}`}
              className={`flex items-center justify-between p-4 rounded-lg cursor-pointer transition-all duration-200 border ${
                isHovered
                  ? 'bg-gray-50 dark:bg-gray-700 border-gray-300 dark:border-gray-600 shadow-md'
                  : 'bg-white dark:bg-gray-800 border-gray-200 dark:border-gray-700'
              }`}
              onMouseEnter={() => setHoveredSegment(entry.type)}
              onMouseLeave={() => setHoveredSegment(null)}
            >
              <div className="flex items-center gap-3">
                <div className="flex items-center gap-2">
                  {entry.type === 'ЭДО' ? (
                    <Zap className="w-5 h-5 text-green-500" />
                  ) : (
                    <FileText className="w-5 h-5 text-amber-500" />
                  )}
                  <div
                    className="w-4 h-4 rounded-full shadow-sm"
                    style={{ backgroundColor: entry.color }}
                  />
                </div>
                <div>
                  <div className="font-semibold text-gray-900 dark:text-white">{entry.name}</div>
                  <div className="text-sm text-gray-500 dark:text-gray-400">{entry.description}</div>
                </div>
              </div>
              <div className="text-right">
                <div className="text-lg font-bold text-gray-900 dark:text-white">{entry.count}</div>
                <div className="text-sm text-gray-500 dark:text-gray-400">{entry.percentage}%</div>
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  if (isLoading) {
    return (
      <div className="h-full flex items-center justify-center">
        <div className="flex items-center gap-2 text-gray-500">
          <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-green-500"></div>
          <span>Загрузка данных...</span>
        </div>
      </div>
    );
  }

  const edoData = data.find(item => item.type === 'ЭДО');

  return (
    <div className="h-full flex flex-col">
      {/* Header with key metrics */}
      <div className="mb-4">
        <div className="grid grid-cols-3 gap-4 mb-4">
          <div className="text-center p-3 bg-gradient-to-r from-blue-50 to-cyan-50 dark:from-blue-900/20 dark:to-cyan-900/20 rounded-lg">
            <div className="flex items-center justify-center gap-1 mb-1">
              <Archive className="w-4 h-4 text-blue-600 dark:text-blue-400" />
              <span className="text-xs text-blue-600 dark:text-blue-400 font-medium">Всего</span>
            </div>
            <div className="text-lg font-bold text-blue-900 dark:text-blue-100">{totalDocuments}</div>
            <div className="text-xs text-gray-500 dark:text-gray-400">документов</div>
          </div>
          
          <div className="text-center p-3 bg-gradient-to-r from-green-50 to-emerald-50 dark:from-green-900/20 dark:to-emerald-900/20 rounded-lg">
            <div className="flex items-center justify-center gap-1 mb-1">
              <Zap className="w-4 h-4 text-green-600 dark:text-green-400" />
              <span className="text-xs text-green-600 dark:text-green-400 font-medium">ЭДО</span>
            </div>
            <div className="text-lg font-bold text-green-900 dark:text-green-100">{edoData?.count || 0}</div>
            <div className="text-xs text-gray-500 dark:text-gray-400">электронных</div>
          </div>
          
          <div className="text-center p-3 bg-gradient-to-r from-amber-50 to-orange-50 dark:from-amber-900/20 dark:to-orange-900/20 rounded-lg">
            <div className="flex items-center justify-center gap-1 mb-1">
              <TrendingUp className="w-4 h-4 text-amber-600 dark:text-amber-400" />
              <span className="text-xs text-amber-600 dark:text-amber-400 font-medium">Цифровизация</span>
            </div>
            <div className="text-lg font-bold text-amber-900 dark:text-amber-100">{edoEfficiency}%</div>
            <div className="text-xs text-gray-500 dark:text-gray-400">эффективность</div>
          </div>
        </div>
      </div>

      {/* Chart or Empty State */}
      <div className="flex-1 min-h-0 flex flex-col">
        {data.length === 0 ? (
          <div className="flex-1 flex flex-col items-center justify-center bg-gray-50 dark:bg-gray-800/50 rounded-lg p-8">
            <AlertTriangle className="w-12 h-12 text-yellow-500 mb-4" />
            <p className="text-gray-500 dark:text-gray-400 text-center text-lg">Нет данных о документах</p>
            <p className="text-gray-400 dark:text-gray-500 text-sm mt-2 text-center">
              Данные появятся после подключения к SAP OData
            </p>
          </div>
        ) : (
          <>
            <div className="flex-1" style={{ minHeight: '200px' }}>
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={data}
                    cx="50%"
                    cy="50%"
                    labelLine={false}
                    label={renderCustomizedLabel}
                    outerRadius={50}
                    innerRadius={20}
                    paddingAngle={2}
                    dataKey="count"
                    animationBegin={0}
                    animationDuration={800}
                  >
                    {data.map((entry, index) => (
                      <Cell
                        key={`cell-${index}`}
                        fill={entry.color}
                        stroke={hoveredSegment === entry.type ? '#374151' : 'rgba(255,255,255,0.3)'}
                        strokeWidth={hoveredSegment === entry.type ? 3 : 1}
                      />
                    ))}
                  </Pie>
                  <Tooltip content={<CustomTooltip />} />
                </PieChart>
              </ResponsiveContainer>
            </div>

            <div className="mt-2" style={{ maxHeight: '120px', overflowY: 'auto' }}>
              <CustomLegend />
            </div>
          </>
        )}
      </div>

      {/* Progress indicator */}
      <div className="mt-4 pt-4 border-t border-gray-200 dark:border-gray-600">
        <div className="flex items-center justify-between mb-2">
          <span className="text-sm font-medium text-gray-700 dark:text-gray-300">
            Прогресс цифровизации
          </span>
          <span className="text-sm font-bold text-green-600 dark:text-green-400">
            {edoEfficiency}%
          </span>
        </div>
        <div className="w-full bg-gray-200 dark:bg-gray-700 rounded-full h-2">
          <div
            className="bg-gradient-to-r from-green-500 to-emerald-500 h-2 rounded-full transition-all duration-1000"
            style={{ width: `${edoEfficiency}%` }}
          />
        </div>
        <div className="flex justify-between mt-1 text-xs text-gray-500 dark:text-gray-400">
          <span>0%</span>
          <span>100%</span>
        </div>
      </div>

      {/* Button for details */}
      <div className="mt-6">
        <button
          onClick={() => setShowDetails(!showDetails)}
          disabled={documentDetails.length === 0}
          className="px-4 py-2 bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed text-white rounded-lg transition-colors"
        >
          {showDetails ? 'Скрыть детали' : 'Показать детальную информацию'}
        </button>
      </div>

      {/* Details table */}
      {showDetails && documentDetails.length > 0 && (
        <div className="mt-6 bg-gray-50 dark:bg-gray-800 rounded-lg p-6">
          <h4 className="text-lg font-semibold text-gray-900 dark:text-white mb-4">
            Детальная информация о документах
          </h4>
          <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-gray-200 dark:divide-gray-600">
              <thead className="bg-gray-100 dark:bg-gray-700">
                <tr>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">ID Документа</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Номер заказа</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Поставщик</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Тип документа</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Статус</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase">Дата</th>
                </tr>
              </thead>
              <tbody className="bg-white dark:bg-gray-800 divide-y divide-gray-200 dark:divide-gray-600">
                {documentDetails.map((doc, index) => (
                  <tr key={index} className="hover:bg-gray-50 dark:hover:bg-gray-700">
                    <td className="px-6 py-4 text-sm font-medium text-gray-900 dark:text-white">{doc.documentId}</td>
                    <td className="px-6 py-4 text-sm text-gray-500 dark:text-gray-300">{doc.orderNumber}</td>
                    <td className="px-6 py-4 text-sm text-gray-500 dark:text-gray-300">{doc.supplierName}</td>
                    <td className="px-6 py-4 text-sm text-gray-500 dark:text-gray-300">{doc.documentType}</td>
                    <td className="px-6 py-4">
                      <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${
                        doc.status === 'Подписан ЭДО' ? 'bg-green-100 text-green-800' :
                        doc.status === 'На согласовании' ? 'bg-yellow-100 text-yellow-800' :
                        doc.status === 'Подписан вручную' ? 'bg-blue-100 text-blue-800' :
                        'bg-red-100 text-red-800'
                      }`}>
                        {doc.status}
                      </span>
                    </td>
                    <td className="px-6 py-4 text-sm text-gray-500 dark:text-gray-300">
                      {doc.date ? new Date(doc.date).toLocaleDateString('ru-RU') : '-'}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
