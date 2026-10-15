import { createContext, useContext, useState, useEffect, type ReactNode } from "react";
import api from "../api/axios";

interface User {
  id: number;
  email: string;
  full_name: string | null;
  avatar_url: string | null;
  role: string;
  is_active: boolean;
}

interface AuthContextType {
  user: User | null;
  token: string | null;
  login: (email: string, password: string) => Promise<void>;
  register: (email: string, password: string, fullName: string, securityQuestion?: string, securityAnswer?: string) => Promise<void>;
  logout: () => void;
  refreshUser: () => Promise<void>;
  isAuthenticated: boolean;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export function AuthProvider({ children }: { children: ReactNode }) {
  const [user, setUser] = useState<User | null>(() => {
    const stored = localStorage.getItem("user");
    return stored ? JSON.parse(stored) : null;
  });
  const [token, setToken] = useState<string | null>(() =>
    localStorage.getItem("token")
  );

  useEffect(() => {
    if (token && !user) {
      api
        .get("/api/auth/me")
        .then((res) => {
          setUser(res.data);
          localStorage.setItem("user", JSON.stringify(res.data));
        })
        .catch(() => {
          setToken(null);
          localStorage.removeItem("token");
        });
    }
  }, [token, user]);

  const login = async (email: string, password: string) => {
    const res = await api.post("/api/auth/login", { email, password });
    const accessToken = res.data.access_token;
    setToken(accessToken);
    localStorage.setItem("token", accessToken);

    const meRes = await api.get("/api/auth/me", {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    setUser(meRes.data);
    localStorage.setItem("user", JSON.stringify(meRes.data));
  };

  const register = async (email: string, password: string, fullName: string, securityQuestion?: string, securityAnswer?: string) => {
    await api.post("/api/auth/register", {
      email,
      password,
      full_name: fullName,
      security_question: securityQuestion,
      security_answer: securityAnswer,
    });
  };

  const logout = () => {
    setUser(null);
    setToken(null);
    localStorage.removeItem("token");
    localStorage.removeItem("user");
  };

  const refreshUser = async () => {
    try {
      const res = await api.get("/api/auth/me");
      setUser(res.data);
      localStorage.setItem("user", JSON.stringify(res.data));
    } catch {
      // ignore
    }
  };

  return (
    <AuthContext.Provider
      value={{ user, token, login, register, logout, refreshUser, isAuthenticated: !!token }}
    >
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  const context = useContext(AuthContext);
  if (!context) throw new Error("useAuth must be used within AuthProvider");
  return context;
}
