import React from 'react';
import { useAuth } from '../contexts/AuthContext';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from './ui/card';
import { Button } from '@/components/ui/button';
import { Badge } from './ui/badge';
import { Avatar, AvatarFallback, AvatarImage } from './ui/avatar';
import { 
  User, 
  LogOut, 
  Settings, 
  Shield, 
  Users, 
  Star,
  Building2,
  Crown,
  UserCheck,
  Briefcase
} from 'lucide-react';
// v2.0.0 (SSO): роли и permissions — из ответа бэка /auth/me
// (user.roles / user.permissions). Раньше компонент ждал USER_ROLES из
// AuthContext, которого там не было, и падал при открытии.

// Роли приложения, которые бэк выдаёт по <RoleMapping> (sso-config.xml).
// Неизвестная роль показывается как есть.
const ROLE_META = {
  DIRECTOR:     { label: 'Директор',            icon: Crown,     color: 'bg-gradient-to-r from-purple-500 to-pink-500'   },
  CENTER_HEAD:  { label: 'Руководитель центра', icon: Building2, color: 'bg-gradient-to-r from-blue-500 to-indigo-500'   },
  GROUP_LEADER: { label: 'Руководитель группы', icon: Users,     color: 'bg-gradient-to-r from-green-500 to-emerald-500' },
  SPECIALIST:   { label: 'Специалист',          icon: UserCheck, color: 'bg-gradient-to-r from-orange-500 to-red-500'    },
  PURCHASER:    { label: 'Закупщик',            icon: Briefcase, color: 'bg-gradient-to-r from-gray-500 to-slate-500'    },
};

function roleMeta(role) {
  return ROLE_META[String(role ?? '').toUpperCase()] ?? null;
}

const UserProfile = ({ isDropdown = false, onClose = null }) => {
  const { user, logout } = useAuth();

  if (!user) return null;

  const fullName    = user.fullName || user.username || '';
  const initials    = fullName.split(' ').filter(Boolean).map(n => n[0]).join('').substring(0, 2).toUpperCase() || '?';
  const primaryRole = user.role ?? user.roles?.[0] ?? null;
  const permissions = Array.isArray(user.permissions) ? user.permissions : [];

  const getRoleIcon = (role) => {
    const Icon = roleMeta(role)?.icon ?? User;
    return <Icon className="h-4 w-4" />;
  };

  const getRoleLabel = (role) => roleMeta(role)?.label ?? (role || 'Пользователь');

  const getRoleColor = (role) => roleMeta(role)?.color ?? 'bg-gradient-to-r from-gray-400 to-gray-500';

  const handleLogout = () => {
    logout();
    if (onClose) onClose();
  };

  if (isDropdown) {
    return (
      <Card className="w-80 border-0 shadow-xl bg-white/95 backdrop-blur-sm">
        <CardHeader className="pb-4">
          <div className="flex items-center space-x-3">
            <Avatar className="h-12 w-12 ring-2 ring-blue-500/20">
              <AvatarImage src={user.avatar} alt={fullName} />
              <AvatarFallback className={`text-white ${getRoleColor(primaryRole)}`}>
                {initials}
              </AvatarFallback>
            </Avatar>
            <div className="flex-1 min-w-0">
              <h3 className="font-semibold text-gray-900 truncate">{fullName}</h3>
              <p className="text-sm text-gray-500 truncate">@{user.username}</p>
            </div>
          </div>
        </CardHeader>
        
        <CardContent className="space-y-4">
          <div className="space-y-3">
            <div className="flex items-center justify-between">
              <span className="text-sm text-gray-600">Роль</span>
              <Badge className={`text-white ${getRoleColor(primaryRole)} border-0`}>
                <span className="flex items-center space-x-1">
                  {getRoleIcon(primaryRole)}
                  <span>{getRoleLabel(primaryRole)}</span>
                </span>
              </Badge>
            </div>
            
            {user.group && (
              <div className="flex items-center justify-between">
                <span className="text-sm text-gray-600">Группа</span>
                <Badge variant="outline" className="border-blue-200 text-blue-700">
                  {user.group}
                </Badge>
              </div>
            )}
            
            <div className="flex items-center justify-between">
              <span className="text-sm text-gray-600">Рейтинг</span>
              <div className="flex items-center space-x-1">
                <Star className="h-4 w-4 text-yellow-500 fill-current" />
                <span className="text-sm font-medium text-gray-900">{user.rating}</span>
              </div>
            </div>
          </div>
          
          <div className="border-t pt-4 space-y-2">
            <Button 
              variant="ghost" 
              className="w-full justify-start text-gray-700 hover:text-gray-900 hover:bg-gray-100"
              onClick={() => onClose && onClose()}
            >
              <Settings className="mr-2 h-4 w-4" />
              Настройки
            </Button>
            
            <Button 
              variant="ghost" 
              className="w-full justify-start text-red-600 hover:text-red-700 hover:bg-red-50"
              onClick={handleLogout}
            >
              <LogOut className="mr-2 h-4 w-4" />
              Выйти
            </Button>
          </div>
        </CardContent>
      </Card>
    );
  }

  return (
    <div className="max-w-2xl mx-auto p-6">
      <Card className="border-0 shadow-xl bg-white/95 backdrop-blur-sm">
        <CardHeader className="text-center pb-6">
          <div className="flex flex-col items-center space-y-4">
            <Avatar className="h-24 w-24 ring-4 ring-blue-500/20">
              <AvatarImage src={user.avatar} alt={fullName} />
              <AvatarFallback className={`text-white text-2xl ${getRoleColor(primaryRole)}`}>
                {initials}
              </AvatarFallback>
            </Avatar>
            <div>
              <CardTitle className="text-2xl text-gray-900">{fullName}</CardTitle>
              <CardDescription className="text-lg">@{user.username}</CardDescription>
            </div>
          </div>
        </CardHeader>
        
        <CardContent className="space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-4">
              <h3 className="font-semibold text-gray-900 flex items-center">
                <Shield className="mr-2 h-5 w-5 text-blue-600" />
                Информация о роли
              </h3>
              
              <div className="space-y-3 pl-7">
                <div className="flex items-center justify-between">
                  <span className="text-gray-600">Роль</span>
                  <Badge className={`text-white ${getRoleColor(primaryRole)} border-0`}>
                    <span className="flex items-center space-x-1">
                      {getRoleIcon(primaryRole)}
                      <span>{getRoleLabel(primaryRole)}</span>
                    </span>
                  </Badge>
                </div>
                
                {user.group && (
                  <div className="flex items-center justify-between">
                    <span className="text-gray-600">Группа</span>
                    <Badge variant="outline" className="border-blue-200 text-blue-700">
                      {user.group}
                    </Badge>
                  </div>
                )}
                
                <div className="flex items-center justify-between">
                  <span className="text-gray-600">Рейтинг</span>
                  <div className="flex items-center space-x-1">
                    <Star className="h-4 w-4 text-yellow-500 fill-current" />
                    <span className="font-medium text-gray-900">{user.rating}</span>
                  </div>
                </div>
              </div>
            </div>
            
            <div className="space-y-4">
              <h3 className="font-semibold text-gray-900 flex items-center">
                <User className="mr-2 h-5 w-5 text-green-600" />
                Разрешения
              </h3>
              
              <div className="space-y-2 pl-7">
                {permissions.length === 0 && (
                  <span className="text-sm text-gray-500">Нет прав</span>
                )}
                {permissions.map((permission, index) => (
                  <Badge 
                    key={index} 
                    variant="outline" 
                    className="mr-2 mb-2 border-green-200 text-green-700"
                  >
                    {permission === '*' ? 'Полный доступ' : permission}
                  </Badge>
                ))}
              </div>
            </div>
          </div>
          
          <div className="border-t pt-6 flex justify-center space-x-4">
            <Button variant="outline" className="flex items-center">
              <Settings className="mr-2 h-4 w-4" />
              Настройки профиля
            </Button>
            
            <Button 
              variant="destructive" 
              onClick={handleLogout}
              className="flex items-center"
            >
              <LogOut className="mr-2 h-4 w-4" />
              Выйти из системы
            </Button>
          </div>
        </CardContent>
      </Card>
    </div>
  );
};

export default UserProfile;

