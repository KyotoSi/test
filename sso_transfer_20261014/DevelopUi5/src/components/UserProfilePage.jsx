// src/components/UserProfilePage.jsx
import React, { useState, useEffect, useRef, useCallback } from 'react';
import { useAuth } from '@/contexts/AuthContext';
import { scopedKey } from '@/api/session';
import {
  User,
  Mail,
  Building,
  Clock,
  Edit2,
  Save,
  X,
  Phone,
  Briefcase,
  MapPin,
  RefreshCw,
  Loader2,
  Info,
  Camera,
  Upload
} from 'lucide-react';

// Константы золотого сечения
const PHI = 1.618033988749895;
const PHI_INVERSE = 0.618033988749895;

// Типографика на основе золотого сечения (базовый размер 16px)
const GOLDEN_SCALE = {
  xs: '0.618rem',    // ~10px
  sm: '0.786rem',    // ~12.5px
  base: '1rem',      // 16px
  md: '1.272rem',    // ~20px
  lg: '1.618rem',    // ~26px
  xl: '2.058rem',    // ~33px
  '2xl': '3.295rem', // ~53px
};

// Отступы на основе золотого сечения
const GOLDEN_SPACING = {
  xs: '0.382rem',    // ~6px
  sm: '0.618rem',    // ~10px
  base: '1rem',      // 16px
  md: '1.618rem',    // ~26px
  lg: '2.618rem',    // ~42px
  xl: '4.236rem',    // ~68px
};

const DEFAULT_INNER_GRADIENT = 'linear-gradient(145deg, rgba(251, 146, 60, 0.55) 0%, rgba(234, 88, 12, 0.27) 100%)';

const clamp = (v, min = 0, max = 100) => Math.min(Math.max(v, min), max);
const round = (v, precision = 3) => parseFloat(v.toFixed(precision));
const adjust = (v, fMin, fMax, tMin, tMax) => round(tMin + ((tMax - tMin) * (v - fMin)) / (fMax - fMin));

const ProfileCard3D = ({
  avatarUrl,
  name,
  title,
  handle,
  status,
  company,
  onPhotoClick
}) => {
  const wrapRef = useRef(null);
  const shellRef = useRef(null);
  const [isActive, setIsActive] = useState(false);

  const getOffsets = (evt, el) => {
    const rect = el.getBoundingClientRect();
    return { x: evt.clientX - rect.left, y: evt.clientY - rect.top };
  };

  const setVarsFromXY = (x, y) => {
    const shell = shellRef.current;
    const wrap = wrapRef.current;
    if (!shell || !wrap) return;

    const width = shell.clientWidth || 1;
    const height = shell.clientHeight || 1;

    const percentX = clamp((100 / width) * x);
    const percentY = clamp((100 / height) * y);

    const centerX = percentX - 50;
    const centerY = percentY - 50;

    wrap.style.setProperty('--pointer-x', `${percentX}%`);
    wrap.style.setProperty('--pointer-y', `${percentY}%`);
    wrap.style.setProperty('--background-x', `${adjust(percentX, 0, 100, 35, 65)}%`);
    wrap.style.setProperty('--background-y', `${adjust(percentY, 0, 100, 35, 65)}%`);
    wrap.style.setProperty('--rotate-x', `${round(-(centerX / 5))}deg`);
    wrap.style.setProperty('--rotate-y', `${round(centerY / 4)}deg`);
  };

  const handlePointerMove = useCallback((event) => {
    const shell = shellRef.current;
    if (!shell) return;
    const { x, y } = getOffsets(event, shell);
    setVarsFromXY(x, y);
  }, []);

  const handlePointerEnter = useCallback((event) => {
    const shell = shellRef.current;
    if (!shell) return;
    setIsActive(true);
    const { x, y } = getOffsets(event, shell);
    setVarsFromXY(x, y);
  }, []);

  const handlePointerLeave = useCallback(() => {
    const shell = shellRef.current;
    if (!shell) return;
    setIsActive(false);
    const centerX = shell.clientWidth / 2;
    const centerY = shell.clientHeight / 2;
    setVarsFromXY(centerX, centerY);
  }, []);

  useEffect(() => {
    const shell = shellRef.current;
    if (!shell) return;

    shell.addEventListener('pointerenter', handlePointerEnter);
    shell.addEventListener('pointermove', handlePointerMove);
    shell.addEventListener('pointerleave', handlePointerLeave);

    const centerX = shell.clientWidth / 2;
    const centerY = shell.clientHeight / 2;
    setVarsFromXY(centerX, centerY);

    return () => {
      shell.removeEventListener('pointerenter', handlePointerEnter);
      shell.removeEventListener('pointermove', handlePointerMove);
      shell.removeEventListener('pointerleave', handlePointerLeave);
    };
  }, [handlePointerEnter, handlePointerMove, handlePointerLeave]);

  return (
    <div
      ref={wrapRef}
      className="profile-card-wrapper"
      style={{
        '--inner-gradient': DEFAULT_INNER_GRADIENT,
        '--rotate-x': '0deg',
        '--rotate-y': '0deg',
        '--pointer-x': '50%',
        '--pointer-y': '50%',
        '--background-x': '50%',
        '--background-y': '50%'
      }}
    >
      <style>{`
        .profile-card-wrapper {
          perspective: 1000px;
          width: 100%;
          height: 100%;
        }

        .profile-card-shell {
          position: relative;
          width: 100%;
          height: 100%;
          transform-style: preserve-3d;
          transition: transform 0.1s ease-out;
          transform:
            rotateY(var(--rotate-x, 0deg))
            rotateX(var(--rotate-y, 0deg));
        }

        .profile-card-shell.active {
          transition: transform 0.05s ease-out;
        }

        .profile-card {
          position: relative;
          width: 100%;
          height: 100%;
          border-radius: ${GOLDEN_SPACING.md};
          overflow: hidden;
          background: linear-gradient(145deg, #fff 0%, #fef3e7 100%);
          border: 2px solid rgba(251, 146, 60, 0.2);
          box-shadow:
            0 ${GOLDEN_SPACING.md} ${GOLDEN_SPACING.xl} rgba(251, 146, 60, 0.15),
            0 ${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.lg} rgba(0, 0, 0, 0.08);
        }

        .profile-card-inside {
          position: relative;
          width: 100%;
          height: 100%;
          z-index: 2;
        }

        .profile-card-shine {
          position: absolute;
          inset: 0;
          background:
            radial-gradient(
              circle at var(--pointer-x, 50%) var(--pointer-y, 50%),
              rgba(255, 255, 255, 0.3) 0%,
              transparent 61.8%
            );
          opacity: 0;
          transition: opacity 0.3s ease;
          pointer-events: none;
          z-index: 10;
        }

        .profile-card-shell.active .profile-card-shine {
          opacity: 1;
        }

        .profile-card-glare {
          position: absolute;
          inset: 0;
          background:
            linear-gradient(
              115deg,
              transparent 38.2%,
              rgba(255, 255, 255, 0.5) 50%,
              transparent 61.8%
            );
          background-position: var(--background-x, 50%) var(--background-y, 50%);
          background-size: 200% 200%;
          opacity: 0;
          transition: opacity 0.3s ease;
          pointer-events: none;
          z-index: 9;
          mix-blend-mode: overlay;
        }

        .profile-card-shell.active .profile-card-glare {
          opacity: 0.618;
        }

        .profile-card-content {
          position: relative;
          padding: ${GOLDEN_SPACING.lg};
          z-index: 3;
          height: 100%;
          display: flex;
          flex-direction: column;
        }

        .profile-avatar-section {
          display: flex;
          flex-direction: column;
          align-items: center;
          margin-bottom: ${GOLDEN_SPACING.md};
        }

        .profile-avatar-wrapper {
          position: relative;
          margin-bottom: ${GOLDEN_SPACING.base};
        }

        .profile-avatar {
          width: calc(${GOLDEN_SPACING.xl} * 1.618);
          height: calc(${GOLDEN_SPACING.xl} * 1.618);
          border-radius: 50%;
          border: 3px solid rgba(251, 146, 60, 0.25);
          box-shadow: 0 ${GOLDEN_SPACING.base} ${GOLDEN_SPACING.lg} rgba(251, 146, 60, 0.2);
          object-fit: cover;
          background: linear-gradient(135deg, #fb923c 0%, #ea580c 100%);
        }

        .profile-avatar-placeholder {
          width: calc(${GOLDEN_SPACING.xl} * 1.618);
          height: calc(${GOLDEN_SPACING.xl} * 1.618);
          border-radius: 50%;
          border: 3px solid rgba(251, 146, 60, 0.25);
          background: linear-gradient(135deg, #fb923c 0%, #ea580c 100%);
          display: flex;
          align-items: center;
          justify-content: center;
          box-shadow: 0 ${GOLDEN_SPACING.base} ${GOLDEN_SPACING.lg} rgba(251, 146, 60, 0.2);
        }

        .profile-camera-btn {
          position: absolute;
          bottom: 0;
          right: 0;
          width: ${GOLDEN_SPACING.lg};
          height: ${GOLDEN_SPACING.lg};
          border-radius: 50%;
          background: linear-gradient(135deg, #fb923c 0%, #ea580c 100%);
          border: 2px solid white;
          display: flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
          box-shadow: 0 ${GOLDEN_SPACING.xs} ${GOLDEN_SPACING.base} rgba(251, 146, 60, 0.3);
        }

        .profile-camera-btn:hover {
          transform: scale(1.1);
          box-shadow: 0 ${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.md} rgba(251, 146, 60, 0.4);
        }

        .profile-info {
          text-align: center;
          flex: 1;
        }

        .profile-name {
          font-size: ${GOLDEN_SCALE.lg};
          font-weight: 700;
          color: #1f2937;
          margin-bottom: ${GOLDEN_SPACING.xs};
          line-height: 1.618;
        }

        .profile-title {
          font-size: ${GOLDEN_SCALE.sm};
          color: #fb923c;
          font-weight: 600;
          margin-bottom: ${GOLDEN_SPACING.base};
          line-height: 1.618;
        }

        .profile-details-grid {
          display: flex;
          flex-direction: column;
          gap: ${GOLDEN_SPACING.sm};
          margin-top: ${GOLDEN_SPACING.base};
        }

        .profile-detail-item {
          display: flex;
          align-items: center;
          gap: ${GOLDEN_SPACING.sm};
          padding: ${GOLDEN_SPACING.sm};
          background: rgba(251, 146, 60, 0.04);
          border-radius: ${GOLDEN_SPACING.base};
          border: 1px solid rgba(251, 146, 60, 0.12);
        }

        .profile-detail-icon {
          flex-shrink: 0;
          width: ${GOLDEN_SPACING.lg};
          height: ${GOLDEN_SPACING.lg};
          display: flex;
          align-items: center;
          justify-content: center;
          background: linear-gradient(135deg, #fb923c 0%, #ea580c 100%);
          border-radius: ${GOLDEN_SPACING.sm};
          color: white;
        }

        .profile-detail-text {
          font-size: ${GOLDEN_SCALE.sm};
          color: #4b5563;
          font-weight: 500;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
          line-height: 1.618;
        }

        .profile-footer {
          margin-top: auto;
          padding-top: ${GOLDEN_SPACING.base};
          border-top: 2px solid rgba(251, 146, 60, 0.12);
          display: flex;
          justify-content: space-between;
          align-items: center;
        }

        .profile-handle {
          font-size: ${GOLDEN_SCALE.sm};
          color: #9ca3af;
          font-weight: 500;
          line-height: 1.618;
        }

        .profile-status {
          display: flex;
          align-items: center;
          gap: ${GOLDEN_SPACING.xs};
          font-size: ${GOLDEN_SCALE.sm};
          color: #10b981;
          font-weight: 600;
          line-height: 1.618;
        }

        .profile-status-dot {
          width: ${GOLDEN_SPACING.xs};
          height: ${GOLDEN_SPACING.xs};
          border-radius: 50%;
          background: #10b981;
          animation: pulse 2s ease-in-out infinite;
        }

        @keyframes pulse {
          0%, 100% { opacity: 1; }
          50% { opacity: 0.618; }
        }
      `}</style>

      <div ref={shellRef} className={`profile-card-shell ${isActive ? 'active' : ''}`}>
        <div className="profile-card">
          <div className="profile-card-inside">
            <div className="profile-card-shine" />
            <div className="profile-card-glare" />
            
            <div className="profile-card-content">
              <div className="profile-avatar-section">
                <div className="profile-avatar-wrapper">
                  {avatarUrl ? (
                    <img
                      src={avatarUrl}
                      alt={name}
                      className="profile-avatar"
                    />
                  ) : (
                    <div className="profile-avatar-placeholder">
                      <User size={42} className="text-white" />
                    </div>
                  )}
                  <button
                    className="profile-camera-btn"
                    onClick={onPhotoClick}
                    type="button"
                  >
                    <Camera size={16} className="text-white" />
                  </button>
                </div>
              </div>

              <div className="profile-info">
                <h3 className="profile-name">{name}</h3>
                <p className="profile-title">{title}</p>

                <div className="profile-details-grid">
                  <div className="profile-detail-item">
                    <div className="profile-detail-icon">
                      <Building size={16} />
                    </div>
                    <span className="profile-detail-text">{company}</span>
                  </div>
                  
                  <div className="profile-detail-item">
                    <div className="profile-detail-icon">
                      <User size={16} />
                    </div>
                    <span className="profile-detail-text">@{handle}</span>
                  </div>
                </div>
              </div>

              <div className="profile-footer">
                <span className="profile-handle">ID: {handle}</span>
                <div className="profile-status">
                  <span className="profile-status-dot" />
                  {status}
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

const UserProfilePage = () => {
  const { user, refreshUser, loading } = useAuth();
  const [isEditing, setIsEditing] = useState(false);
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [userPhoto, setUserPhoto] = useState(null);
  const [showPhotoUpload, setShowPhotoUpload] = useState(false);

  const [profileData, setProfileData] = useState({
    username: '',
    firstName: '',
    lastName: '',
    fullName: '',
    email: '',
    phone: '',
    department: '',
    company: '',
    city: '',
    position: '',
  });

  useEffect(() => {
    // Фото привязано к пользователю SAP (sessionStorage чистится при logout)
    const key = scopedKey('userPhoto');
    const savedPhoto = key ? sessionStorage.getItem(key) : null;
    if (savedPhoto) {
      setUserPhoto(savedPhoto);
    }
  }, []);

  useEffect(() => {
    if (user) {
      setProfileData({
        username: user.username || user.bname || '',
        firstName: user.firstName || '',
        lastName: user.lastName || '',
        fullName: user.fullName || '',
        email: user.email || '',
        phone: user.phone || '',
        department: user.department || '',
        company: user.company || '',
        city: user.city || '',
        position: user.position || '',
      });
    }
  }, [user]);

  const handlePhotoUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        const photoData = reader.result;
        setUserPhoto(photoData);
        const key = scopedKey('userPhoto');
        if (key) {
          try { sessionStorage.setItem(key, photoData); } catch { /* переполнение — без кэша */ }
        }
        setShowPhotoUpload(false);
      };
      reader.readAsDataURL(file);
    }
  };

  const handleRefresh = async () => {
    setIsRefreshing(true);
    await refreshUser();
    setIsRefreshing(false);
  };

  const handleEdit = () => setIsEditing(true);

  const handleSave = () => {
    setIsEditing(false);
    console.log('💾 Сохранение профиля:', profileData);
  };

  const handleCancel = () => {
    setIsEditing(false);
    if (user) {
      setProfileData({
        username: user.username || user.bname || '',
        firstName: user.firstName || '',
        lastName: user.lastName || '',
        fullName: user.fullName || '',
        email: user.email || '',
        phone: user.phone || '',
        department: user.department || '',
        company: user.company || '',
        city: user.city || '',
        position: user.position || '',
      });
    }
  };

  const handleChange = (field, value) => {
    setProfileData(prev => ({ ...prev, [field]: value }));
  };

  const formatDate = (timestamp) => {
    if (!timestamp) return 'Неизвестно';
    try {
      const date = new Date(timestamp);
      return date.toLocaleString('ru-RU', {
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
    } catch {
      return 'Неизвестно';
    }
  };

  if (loading && !user) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-orange-50 via-yellow-50 to-amber-50 flex items-center justify-center">
        <div className="text-center">
          <Loader2 className="w-12 h-12 animate-spin text-orange-500 mx-auto mb-4" />
          <p className="text-gray-600 font-medium" style={{ fontSize: GOLDEN_SCALE.base }}>
            Загрузка профиля...
          </p>
        </div>
      </div>
    );
  }

  return (
    <div 
      className="min-h-screen bg-gradient-to-br from-orange-50 via-yellow-50 to-amber-50"
      style={{ padding: GOLDEN_SPACING.lg }}
    >
      <div className="container mx-auto" style={{ maxWidth: '1400px' }}>
        {/* Основной контейнер с золотым соотношением сторон */}
        <div 
          className="bg-white rounded-3xl shadow-2xl overflow-hidden border border-orange-100"
          style={{ borderRadius: GOLDEN_SPACING.md }}
        >
          <div style={{ padding: `${GOLDEN_SPACING.lg} ${GOLDEN_SPACING.xl}` }}>
            
            {/* Заголовок с пропорциями золотого сечения */}
            <div 
              className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4"
              style={{ marginBottom: GOLDEN_SPACING.xl }}
            >
              <div style={{ flex: PHI }}>
                <h1 
                  className="font-bold text-gray-800"
                  style={{ 
                    fontSize: GOLDEN_SCALE.xl,
                    marginBottom: GOLDEN_SPACING.sm,
                    lineHeight: 1.618
                  }}
                >
                  Профиль пользователя
                </h1>
                <p 
                  className="text-gray-600"
                  style={{ 
                    fontSize: GOLDEN_SCALE.sm,
                    lineHeight: 1.618
                  }}
                >
                  Управление личными данными и настройками
                </p>
              </div>

              <div 
                className="flex gap-3"
                style={{ flex: 1, justifyContent: 'flex-end' }}
              >
                <button
                  onClick={handleRefresh}
                  disabled={isRefreshing}
                  className="flex items-center gap-2 bg-white border-2 border-orange-300 text-orange-600 rounded-xl transition-all shadow-md hover:shadow-lg transform hover:-translate-y-0.5 disabled:opacity-50"
                  style={{
                    padding: `${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.md}`,
                    fontSize: GOLDEN_SCALE.sm,
                    borderRadius: GOLDEN_SPACING.base
                  }}
                >
                  <RefreshCw size={18} className={isRefreshing ? 'animate-spin' : ''} />
                  <span className="font-medium">Обновить</span>
                </button>

                {!isEditing ? (
                  <button
                    onClick={handleEdit}
                    className="flex items-center gap-2 bg-gradient-to-r from-yellow-500 to-orange-500 hover:from-yellow-600 hover:to-orange-600 text-white rounded-xl transition-all shadow-lg hover:shadow-xl transform hover:-translate-y-0.5"
                    style={{
                      padding: `${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.md}`,
                      fontSize: GOLDEN_SCALE.sm,
                      borderRadius: GOLDEN_SPACING.base
                    }}
                  >
                    <Edit2 size={18} />
                    <span className="font-medium">Редактировать</span>
                  </button>
                ) : (
                  <>
                    <button
                      onClick={handleSave}
                      className="flex items-center gap-2 bg-gradient-to-r from-green-500 to-emerald-500 hover:from-green-600 hover:to-emerald-600 text-white rounded-xl transition-all shadow-lg hover:shadow-xl transform hover:-translate-y-0.5"
                      style={{
                        padding: `${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.md}`,
                        fontSize: GOLDEN_SCALE.sm,
                        borderRadius: GOLDEN_SPACING.base
                      }}
                    >
                      <Save size={18} />
                      <span className="font-medium">Сохранить</span>
                    </button>
                    <button
                      onClick={handleCancel}
                      className="flex items-center gap-2 bg-gradient-to-r from-gray-500 to-gray-600 hover:from-gray-600 hover:to-gray-700 text-white rounded-xl transition-all shadow-lg hover:shadow-xl transform hover:-translate-y-0.5"
                      style={{
                        padding: `${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.md}`,
                        fontSize: GOLDEN_SCALE.sm,
                        borderRadius: GOLDEN_SPACING.base
                      }}
                    >
                      <X size={18} />
                      <span className="font-medium">Отмена</span>
                    </button>
                  </>
                )}
              </div>
            </div>

            {/* Модальное окно загрузки фото */}
            {showPhotoUpload && (
              <div 
                className="bg-orange-50 border-2 border-orange-200"
                style={{
                  marginBottom: GOLDEN_SPACING.md,
                  padding: GOLDEN_SPACING.md,
                  borderRadius: GOLDEN_SPACING.base
                }}
              >
                <p 
                  className="text-gray-700 font-medium"
                  style={{
                    fontSize: GOLDEN_SCALE.sm,
                    marginBottom: GOLDEN_SPACING.sm,
                    lineHeight: 1.618
                  }}
                >
                  📸 Загрузить фотографию:
                </p>
                <div className="space-y-2">
                  <label 
                    className="flex items-center justify-center gap-2 bg-white border-2 border-orange-300 rounded-lg cursor-pointer hover:bg-orange-50 transition-colors"
                    style={{ 
                      padding: `${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.base}`,
                      borderRadius: GOLDEN_SPACING.sm
                    }}
                  >
                    <Upload size={16} className="text-orange-600" />
                    <span 
                      className="font-medium text-orange-700"
                      style={{ fontSize: GOLDEN_SCALE.sm }}
                    >
                      Выбрать файл
                    </span>
                    <input
                      type="file"
                      accept="image/*"
                      onChange={handlePhotoUpload}
                      className="hidden"
                    />
                  </label>
                  <button
                    onClick={() => setShowPhotoUpload(false)}
                    className="w-full text-gray-600 hover:text-gray-800 transition-colors"
                    style={{ 
                      padding: `${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.base}`,
                      fontSize: GOLDEN_SCALE.sm
                    }}
                  >
                    Отмена
                  </button>
                </div>
                <p 
                  className="text-gray-500"
                  style={{
                    fontSize: GOLDEN_SCALE.xs,
                    marginTop: GOLDEN_SPACING.xs,
                    lineHeight: 1.618
                  }}
                >
                  Импорт из SAP OData пока не реализован
                </p>
              </div>
            )}

            {/* Основные карточки с золотым соотношением */}
            <div 
              className="grid grid-cols-1 lg:grid-cols-2 gap-6"
              style={{ 
                marginBottom: GOLDEN_SPACING.md,
                gap: GOLDEN_SPACING.md
              }}
            >
              {/* 3D Карточка профиля - 38.2% от общей ширины */}
              <div style={{ height: `${GOLDEN_SPACING.xl * 6.854}px` }}>
                <ProfileCard3D
                  avatarUrl={userPhoto}
                  name={profileData.fullName || profileData.username}
                  title={profileData.position || 'Должность не указана'}
                  handle={profileData.username}
                  status="Online"
                  company={profileData.company || 'POLYUS'}
                  onPhotoClick={() => setShowPhotoUpload(!showPhotoUpload)}
                />
              </div>

              {/* Рабочая информация - 61.8% от общей ширины */}
              <div 
                className="bg-gradient-to-br from-amber-50 to-orange-50 border border-amber-200 shadow-md"
                style={{
                  borderRadius: GOLDEN_SPACING.md,
                  padding: GOLDEN_SPACING.md
                }}
              >
                <div 
                  className="flex items-center gap-3"
                  style={{ marginBottom: GOLDEN_SPACING.md }}
                >
                  <div 
                    className="rounded-xl bg-gradient-to-br from-amber-400 to-orange-500 flex items-center justify-center shadow-lg"
                    style={{
                      width: GOLDEN_SPACING.xl,
                      height: GOLDEN_SPACING.xl,
                      borderRadius: GOLDEN_SPACING.base
                    }}
                  >
                    <Building size={24} className="text-white" />
                  </div>
                  <h3 
                    className="font-bold text-gray-800"
                    style={{ 
                      fontSize: GOLDEN_SCALE.md,
                      lineHeight: 1.618
                    }}
                  >
                    Рабочая информация
                  </h3>
                </div>

                <div style={{ display: 'flex', flexDirection: 'column', gap: GOLDEN_SPACING.base }}>
                  <div>
                    <label 
                      className="text-gray-600 font-semibold flex items-center gap-2"
                      style={{
                        fontSize: GOLDEN_SCALE.sm,
                        marginBottom: GOLDEN_SPACING.xs,
                        lineHeight: 1.618
                      }}
                    >
                      <Building size={14} />
                      Компания
                    </label>
                    <input
                      type="text"
                      value={profileData.company}
                      disabled
                      className="w-full bg-white border-2 border-amber-200 text-gray-800 font-medium cursor-not-allowed"
                      style={{
                        padding: `${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.base}`,
                        borderRadius: GOLDEN_SPACING.base,
                        fontSize: GOLDEN_SCALE.sm,
                        lineHeight: 1.618
                      }}
                    />
                  </div>

                  <div>
                    <label 
                      className="text-gray-600 font-semibold flex items-center gap-2"
                      style={{
                        fontSize: GOLDEN_SCALE.sm,
                        marginBottom: GOLDEN_SPACING.xs,
                        lineHeight: 1.618
                      }}
                    >
                      <Briefcase size={14} />
                      Отдел
                    </label>
                    <input
                      type="text"
                      value={profileData.department}
                      disabled
                      className="w-full bg-white border-2 border-amber-200 text-gray-800 font-medium cursor-not-allowed"
                      style={{
                        padding: `${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.base}`,
                        borderRadius: GOLDEN_SPACING.base,
                        fontSize: GOLDEN_SCALE.sm,
                        lineHeight: 1.618
                      }}
                    />
                  </div>

                  <div>
                    <label 
                      className="text-gray-600 font-semibold flex items-center gap-2"
                      style={{
                        fontSize: GOLDEN_SCALE.sm,
                        marginBottom: GOLDEN_SPACING.xs,
                        lineHeight: 1.618
                      }}
                    >
                      <Briefcase size={14} />
                      Должность
                    </label>
                    <input
                      type="text"
                      value={profileData.position}
                      onChange={(e) => handleChange('position', e.target.value)}
                      disabled={!isEditing}
                      className="w-full bg-white border-2 border-amber-200 text-gray-800 disabled:bg-gray-50 disabled:cursor-not-allowed focus:ring-2 focus:ring-amber-400 focus:border-transparent transition-all"
                      style={{
                        padding: `${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.base}`,
                        borderRadius: GOLDEN_SPACING.base,
                        fontSize: GOLDEN_SCALE.sm,
                        lineHeight: 1.618
                      }}
                    />
                  </div>

                  <div>
                    <label 
                      className="text-gray-600 font-semibold flex items-center gap-2"
                      style={{
                        fontSize: GOLDEN_SCALE.sm,
                        marginBottom: GOLDEN_SPACING.xs,
                        lineHeight: 1.618
                      }}
                    >
                      <MapPin size={14} />
                      Город
                    </label>
                    <input
                      type="text"
                      value={profileData.city}
                      disabled
                      className="w-full bg-white border-2 border-amber-200 text-gray-800 font-medium cursor-not-allowed"
                      style={{
                        padding: `${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.base}`,
                        borderRadius: GOLDEN_SPACING.base,
                        fontSize: GOLDEN_SCALE.sm,
                        lineHeight: 1.618
                      }}
                    />
                  </div>

                  <div>
                    <label 
                      className="text-gray-600 font-semibold flex items-center gap-2"
                      style={{
                        fontSize: GOLDEN_SCALE.sm,
                        marginBottom: GOLDEN_SPACING.xs,
                        lineHeight: 1.618
                      }}
                    >
                      <Mail size={14} />
                      Email
                    </label>
                    <input
                      type="email"
                      value={profileData.email}
                      onChange={(e) => handleChange('email', e.target.value)}
                      disabled={!isEditing}
                      className="w-full bg-white border-2 border-amber-200 text-gray-800 disabled:bg-gray-50 disabled:cursor-not-allowed focus:ring-2 focus:ring-amber-400 focus:border-transparent transition-all"
                      style={{
                        padding: `${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.base}`,
                        borderRadius: GOLDEN_SPACING.base,
                        fontSize: GOLDEN_SCALE.sm,
                        lineHeight: 1.618
                      }}
                    />
                  </div>

                  <div>
                    <label 
                      className="text-gray-600 font-semibold flex items-center gap-2"
                      style={{
                        fontSize: GOLDEN_SCALE.sm,
                        marginBottom: GOLDEN_SPACING.xs,
                        lineHeight: 1.618
                      }}
                    >
                      <Phone size={14} />
                      Телефон
                    </label>
                    <input
                      type="tel"
                      value={profileData.phone}
                      onChange={(e) => handleChange('phone', e.target.value)}
                      disabled={!isEditing}
                      placeholder="+7 (___) ___-__-__"
                      className="w-full bg-white border-2 border-amber-200 text-gray-800 disabled:bg-gray-50 disabled:cursor-not-allowed focus:ring-2 focus:ring-amber-400 focus:border-transparent transition-all"
                      style={{
                        padding: `${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.base}`,
                        borderRadius: GOLDEN_SPACING.base,
                        fontSize: GOLDEN_SCALE.sm,
                        lineHeight: 1.618
                      }}
                    />
                  </div>
                </div>
              </div>
            </div>

            {/* Информация о сессии */}
            <div 
              className="bg-gradient-to-br from-blue-50 to-indigo-50 border border-blue-200 shadow-md"
              style={{
                borderRadius: GOLDEN_SPACING.md,
                padding: GOLDEN_SPACING.md
              }}
            >
              <div 
                className="flex items-center gap-3"
                style={{ marginBottom: GOLDEN_SPACING.md }}
              >
                <div 
                  className="rounded-xl bg-gradient-to-br from-blue-400 to-indigo-500 flex items-center justify-center shadow-lg"
                  style={{
                    width: GOLDEN_SPACING.xl,
                    height: GOLDEN_SPACING.xl,
                    borderRadius: GOLDEN_SPACING.base
                  }}
                >
                  <Clock size={24} className="text-white" />
                </div>
                <h3 
                  className="font-bold text-gray-800"
                  style={{ 
                    fontSize: GOLDEN_SCALE.md,
                    lineHeight: 1.618
                  }}
                >
                  Информация о сессии
                </h3>
              </div>

              <div 
                className="grid grid-cols-1 md:grid-cols-3"
                style={{ gap: GOLDEN_SPACING.md }}
              >
                <div>
                  <span 
                    className="text-gray-600 font-semibold block"
                    style={{
                      fontSize: GOLDEN_SCALE.sm,
                      marginBottom: GOLDEN_SPACING.xs,
                      lineHeight: 1.618
                    }}
                  >
                    Статус авторизации:
                  </span>
                  <div>
                    {user?.authenticated ? (
                      <span 
                        className="inline-flex items-center rounded-xl font-semibold bg-green-100 text-green-800 border border-green-300"
                        style={{
                          padding: `${GOLDEN_SPACING.xs} ${GOLDEN_SPACING.base}`,
                          fontSize: GOLDEN_SCALE.sm,
                          borderRadius: GOLDEN_SPACING.base
                        }}
                      >
                        <span 
                          className="bg-green-500 rounded-full mr-2 animate-pulse"
                          style={{
                            width: GOLDEN_SPACING.xs,
                            height: GOLDEN_SPACING.xs
                          }}
                        />
                        Авторизован
                      </span>
                    ) : (
                      <span 
                        className="inline-flex items-center rounded-xl font-semibold bg-red-100 text-red-800 border border-red-300"
                        style={{
                          padding: `${GOLDEN_SPACING.xs} ${GOLDEN_SPACING.base}`,
                          fontSize: GOLDEN_SCALE.sm,
                          borderRadius: GOLDEN_SPACING.base
                        }}
                      >
                        <span 
                          className="bg-red-500 rounded-full mr-2"
                          style={{
                            width: GOLDEN_SPACING.xs,
                            height: GOLDEN_SPACING.xs
                          }}
                        />
                        Не авторизован
                      </span>
                    )}
                  </div>
                </div>

                <div>
                  <span 
                    className="text-gray-600 font-semibold block"
                    style={{
                      fontSize: GOLDEN_SCALE.sm,
                      marginBottom: GOLDEN_SPACING.xs,
                      lineHeight: 1.618
                    }}
                  >
                    SAP Client:
                  </span>
                  <p 
                    className="text-gray-800 font-mono font-semibold bg-white border border-blue-200"
                    style={{
                      padding: `${GOLDEN_SPACING.xs} ${GOLDEN_SPACING.base}`,
                      borderRadius: GOLDEN_SPACING.base,
                      fontSize: GOLDEN_SCALE.sm,
                      lineHeight: 1.618
                    }}
                  >
                    {user?.sapClient || '—'}
                  </p>
                </div>

                <div>
                  <span 
                    className="text-gray-600 font-semibold block"
                    style={{
                      fontSize: GOLDEN_SCALE.sm,
                      marginBottom: GOLDEN_SPACING.xs,
                      lineHeight: 1.618
                    }}
                  >
                    Последний вход:
                  </span>
                  <p 
                    className="text-gray-800 font-medium bg-white border border-blue-200"
                    style={{
                      padding: `${GOLDEN_SPACING.xs} ${GOLDEN_SPACING.base}`,
                      borderRadius: GOLDEN_SPACING.base,
                      fontSize: GOLDEN_SCALE.sm,
                      lineHeight: 1.618
                    }}
                  >
                    {formatDate(user?.timestamp)}
                  </p>
                </div>
              </div>
            </div>

            {/* Примечание */}
            <div 
              className="bg-gradient-to-r from-yellow-50 to-amber-50 border-2 border-yellow-300"
              style={{
                marginTop: GOLDEN_SPACING.md,
                padding: GOLDEN_SPACING.md,
                borderRadius: GOLDEN_SPACING.md
              }}
            >
              <div className="flex items-start gap-3">
                <div 
                  className="rounded-lg bg-yellow-400 flex items-center justify-center flex-shrink-0"
                  style={{
                    width: GOLDEN_SPACING.lg,
                    height: GOLDEN_SPACING.lg,
                    borderRadius: GOLDEN_SPACING.sm,
                    marginTop: GOLDEN_SPACING.xs
                  }}
                >
                  <Info size={18} className="text-white" />
                </div>
                <div>
                  <p 
                    className="text-gray-700 leading-relaxed"
                    style={{
                      fontSize: GOLDEN_SCALE.sm,
                      lineHeight: 1.618
                    }}
                  >
                    <strong className="text-gray-900">Данные из SAP:</strong> Информация загружается из EntitySet <code className="bg-yellow-100 rounded" style={{ padding: `${GOLDEN_SPACING.xs} ${GOLDEN_SPACING.sm}`, fontSize: GOLDEN_SCALE.xs }}>UserInfoSet</code> сервиса <code className="bg-yellow-100 rounded" style={{ padding: `${GOLDEN_SPACING.xs} ${GOLDEN_SPACING.sm}`, fontSize: GOLDEN_SCALE.xs }}>ZMM_MFC_TASK_SRV</code>.
                    Поля с серым фоном управляются системой SAP и доступны только для чтения.
                  </p>
                  {user?.note && (
                    <p 
                      className="text-gray-600 bg-yellow-100 rounded"
                      style={{
                        fontSize: GOLDEN_SCALE.xs,
                        marginTop: GOLDEN_SPACING.xs,
                        padding: `${GOLDEN_SPACING.sm} ${GOLDEN_SPACING.base}`,
                        lineHeight: 1.618
                      }}
                    >
                      ⚠️ {user.note}
                    </p>
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default UserProfilePage;
