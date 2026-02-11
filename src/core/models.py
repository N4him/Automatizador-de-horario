"""
Modelos de datos del sistema
"""
from dataclasses import dataclass, field
from typing import List, Dict, Tuple


@dataclass
class Monitor:
    """Representa un monitor con su disponibilidad"""
    id: int
    nombre: str
    prioridad: int = 3
    min: int = 8
    max: int = 20
    horas: int = 0
    disp: Dict[str, List[Tuple[int, int]]] = field(default_factory=dict)
    asignaciones: List[Dict] = field(default_factory=list)
    
    def to_dict(self):
        """Convierte el monitor a diccionario"""
        return {
            "id": self.id,
            "nombre": self.nombre,
            "prioridad": self.prioridad,
            "min": self.min,
            "max": self.max,
            "horas": self.horas,
            "disp": self.disp,
            "asignaciones": self.asignaciones
        }
    
    @classmethod
    def from_dict(cls, data):
        """Crea un monitor desde un diccionario"""
        return cls(**data)


@dataclass
class Espacio:
    """Representa un espacio/horario a asignar"""
    sala: str
    dia: str
    hora_inicio: int
    hora_fin: int
    curso: str
    duracion: int
    monitor: str = None
    prioridad: int = None
    estado: str = "❌"