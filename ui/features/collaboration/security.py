# -*- coding: utf-8 -*-
"""协作安全模块 - 加密、访问控制"""

import base64
import hashlib
import hmac
import os
import re
import secrets
import time
from dataclasses import dataclass, field
from typing import Optional, List, Set, Dict
from enum import Enum

try:
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    HAS_CRYPTO = True
except ImportError:
    HAS_CRYPTO = False


class PasswordStrength(Enum):
    """密码强度等级"""
    WEAK = "weak"
    FAIR = "fair"
    GOOD = "good"
    STRONG = "strong"


@dataclass
class PasswordAnalysis:
    """密码分析结果"""
    strength: PasswordStrength
    score: int  # 0-100
    suggestions: List[str] = field(default_factory=list)
    
    @property
    def strength_text(self) -> str:
        texts = {
            PasswordStrength.WEAK: "弱",
            PasswordStrength.FAIR: "一般",
            PasswordStrength.GOOD: "良好",
            PasswordStrength.STRONG: "强",
        }
        return texts.get(self.strength, "未知")
    
    @property
    def strength_color(self) -> str:
        colors = {
            PasswordStrength.WEAK: "#ef4444",
            PasswordStrength.FAIR: "#f59e0b",
            PasswordStrength.GOOD: "#10b981",
            PasswordStrength.STRONG: "#059669",
        }
        return colors.get(self.strength, "#6b7280")


class PasswordChecker:
    """密码强度检测器"""
    
    @staticmethod
    def analyze(password: str) -> PasswordAnalysis:
        """分析密码强度"""
        score = 0
        suggestions = []
        
        # 长度检查
        length = len(password)
        if length >= 8:
            score += 20
        elif length >= 6:
            score += 10
        else:
            suggestions.append("密码长度至少8位")
        
        if length >= 12:
            score += 10
        
        # 包含小写字母
        if re.search(r'[a-z]', password):
            score += 15
        else:
            suggestions.append("添加小写字母")
        
        # 包含大写字母
        if re.search(r'[A-Z]', password):
            score += 15
        else:
            suggestions.append("添加大写字母")
        
        # 包含数字
        if re.search(r'\d', password):
            score += 15
        else:
            suggestions.append("添加数字")
        
        # 包含特殊字符
        if re.search(r'[!@#$%^&*(),.?":{}|<>]', password):
            score += 20
        else:
            suggestions.append("添加特殊字符")
        
        # 不包含常见模式
        common_patterns = ['123', 'abc', 'qwerty', 'password', '111', '000']
        for pattern in common_patterns:
            if pattern.lower() in password.lower():
                score -= 10
                suggestions.append("避免使用常见模式")
                break
        
        # 确定强度等级
        score = max(0, min(100, score))
        
        if score >= 80:
            strength = PasswordStrength.STRONG
        elif score >= 60:
            strength = PasswordStrength.GOOD
        elif score >= 40:
            strength = PasswordStrength.FAIR
        else:
            strength = PasswordStrength.WEAK
        
        return PasswordAnalysis(strength=strength, score=score, suggestions=suggestions)


class Encryptor:
    """端到端加密器"""
    
    def __init__(self, password: str = None):
        self._key: Optional[bytes] = None
        self._fernet: Optional['Fernet'] = None
        
        if password and HAS_CRYPTO:
            self._derive_key(password)
    
    def _derive_key(self, password: str, salt: bytes = None):
        """从密码派生密钥"""
        if not HAS_CRYPTO:
            return
        
        if salt is None:
            salt = os.urandom(16)
        
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=100000,
        )
        
        key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
        self._key = key
        self._fernet = Fernet(key)
        self._salt = salt
    
    def encrypt(self, data: str) -> str:
        """加密数据"""
        if not self._fernet:
            return data
        
        encrypted = self._fernet.encrypt(data.encode())
        return base64.urlsafe_b64encode(encrypted).decode()
    
    def decrypt(self, encrypted_data: str) -> str:
        """解密数据"""
        if not self._fernet:
            return encrypted_data
        
        try:
            data = base64.urlsafe_b64decode(encrypted_data.encode())
            decrypted = self._fernet.decrypt(data)
            return decrypted.decode()
        except Exception:
            return encrypted_data
    
    @property
    def is_enabled(self) -> bool:
        """是否启用加密"""
        return self._fernet is not None


class AccessControl:
    """访问控制"""
    
    def __init__(self):
        self._is_locked = False
        self._waiting_room_enabled = False
        self._waiting_room: List[Dict] = []
        self._ip_whitelist: Set[str] = set()
        self._ip_blacklist: Set[str] = set()
        self._max_participants = 50
        self._allowed_participants: Set[str] = set()  # 允许的参与者ID
    
    @property
    def is_locked(self) -> bool:
        """会议是否锁定"""
        return self._is_locked
    
    @property
    def waiting_room_enabled(self) -> bool:
        """等候室是否启用"""
        return self._waiting_room_enabled
    
    @property
    def waiting_count(self) -> int:
        """等候室人数"""
        return len(self._waiting_room)
    
    def lock_meeting(self):
        """锁定会议"""
        self._is_locked = True
    
    def unlock_meeting(self):
        """解锁会议"""
        self._is_locked = False
    
    def enable_waiting_room(self, enabled: bool = True):
        """启用/禁用等候室"""
        self._waiting_room_enabled = enabled
    
    def add_to_waiting_room(self, participant: Dict) -> bool:
        """添加到等候室"""
        if not self._waiting_room_enabled:
            return False
        
        self._waiting_room.append({
            **participant,
            'join_time': time.time()
        })
        return True
    
    def get_waiting_list(self) -> List[Dict]:
        """获取等候列表"""
        return list(self._waiting_room)
    
    def admit_participant(self, participant_id: str) -> Optional[Dict]:
        """允许参与者进入"""
        for i, p in enumerate(self._waiting_room):
            if p.get('id') == participant_id:
                self._allowed_participants.add(participant_id)
                return self._waiting_room.pop(i)
        return None
    
    def reject_participant(self, participant_id: str) -> bool:
        """拒绝参与者"""
        for i, p in enumerate(self._waiting_room):
            if p.get('id') == participant_id:
                self._waiting_room.pop(i)
                return True
        return False
    
    def admit_all(self) -> List[Dict]:
        """允许所有等候者进入"""
        admitted = list(self._waiting_room)
        for p in admitted:
            self._allowed_participants.add(p.get('id', ''))
        self._waiting_room = []
        return admitted
    
    def add_ip_whitelist(self, ip: str):
        """添加IP白名单"""
        self._ip_whitelist.add(ip)
    
    def remove_ip_whitelist(self, ip: str):
        """移除IP白名单"""
        self._ip_whitelist.discard(ip)
    
    def add_ip_blacklist(self, ip: str):
        """添加IP黑名单"""
        self._ip_blacklist.add(ip)
    
    def remove_ip_blacklist(self, ip: str):
        """移除IP黑名单"""
        self._ip_blacklist.discard(ip)
    
    def check_ip_access(self, ip: str) -> bool:
        """检查IP访问权限"""
        # 黑名单优先
        if ip in self._ip_blacklist:
            return False
        
        # 如果有白名单，只允许白名单内的IP
        if self._ip_whitelist:
            return ip in self._ip_whitelist
        
        return True
    
    def can_join(self, participant_id: str, ip: str = None) -> tuple[bool, str]:
        """检查是否可以加入"""
        # 检查IP
        if ip and not self.check_ip_access(ip):
            return False, "IP地址被禁止访问"
        
        # 检查锁定状态
        if self._is_locked and participant_id not in self._allowed_participants:
            return False, "会议已锁定"
        
        # 检查等候室
        if self._waiting_room_enabled and participant_id not in self._allowed_participants:
            return False, "请在等候室等待主持人允许"
        
        return True, ""
    
    def set_max_participants(self, max_count: int):
        """设置最大参与者数"""
        self._max_participants = max_count
    
    def get_settings(self) -> Dict:
        """获取访问控制设置"""
        return {
            'is_locked': self._is_locked,
            'waiting_room_enabled': self._waiting_room_enabled,
            'ip_whitelist': list(self._ip_whitelist),
            'ip_blacklist': list(self._ip_blacklist),
            'max_participants': self._max_participants,
        }
    
    def apply_settings(self, settings: Dict):
        """应用访问控制设置"""
        self._is_locked = settings.get('is_locked', False)
        self._waiting_room_enabled = settings.get('waiting_room_enabled', False)
        self._ip_whitelist = set(settings.get('ip_whitelist', []))
        self._ip_blacklist = set(settings.get('ip_blacklist', []))
        self._max_participants = settings.get('max_participants', 50)


class TokenManager:
    """令牌管理器"""
    
    def __init__(self, secret_key: str = None):
        self._secret_key = secret_key or secrets.token_hex(32)
    
    def generate_token(self, participant_id: str, expires_in: int = 3600) -> str:
        """生成访问令牌"""
        payload = f"{participant_id}:{int(time.time()) + expires_in}"
        signature = hmac.new(
            self._secret_key.encode(),
            payload.encode(),
            hashlib.sha256
        ).hexdigest()
        
        token = base64.urlsafe_b64encode(f"{payload}:{signature}".encode()).decode()
        return token
    
    def verify_token(self, token: str) -> tuple[bool, Optional[str]]:
        """验证令牌"""
        try:
            decoded = base64.urlsafe_b64decode(token.encode()).decode()
            parts = decoded.rsplit(':', 2)
            
            if len(parts) != 3:
                return False, None
            
            participant_id, expires_str, signature = parts
            
            # 验证签名
            payload = f"{participant_id}:{expires_str}"
            expected_signature = hmac.new(
                self._secret_key.encode(),
                payload.encode(),
                hashlib.sha256
            ).hexdigest()
            
            if not hmac.compare_digest(signature, expected_signature):
                return False, None
            
            # 验证过期时间
            if int(expires_str) < time.time():
                return False, None
            
            return True, participant_id
            
        except Exception:
            return False, None
