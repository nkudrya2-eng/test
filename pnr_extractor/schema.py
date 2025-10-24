JSON_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "project": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "project_code": {"type": "string", "minLength": 1},
                "project_title": {"type": "string", "minLength": 1},
                "site_title": {"type": "string"},
                "designer": {"type": "string"},
                "customer": {"type": "string"},
                "executor": {"type": "string", "default": "ООО «Азимут»"},
                "category": {
                    "type": "string",
                    "enum": ["I", "II", "III"],
                    "default": "II"
                }
            },
            "required": ["project_code", "project_title"]
        },

        "equipment": {
            "type": "array",
            "minItems": 1,
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "equipment_type": {
                        "type": "string",
                        "enum": [
                            "switch","router","sfp","pdu","ups","voip_phone","din_panel",
                            "ground_busbar","breaker","socket_din","other","cctv","cabinet",
                            "cable_management","access_control","computer","ntp-clock"
                        ]
                    },
                    "priority": {"type": "integer", "minimum": 1, "maximum": 10},
                    "vendor": {"type": "string"},
                    "name": {"type": "string", "minLength": 1},
                    "model_or_code": {"type": "string"},
                    "description": {"type": "string", "maxLength": 180},
                    "quantity": {"type": "integer", "minimum": 1},
                    # Для switch/router — допустим 0 (если не распознано), но не отрицательно
                    "port_count": {"type": "integer", "minimum": 0, "default": 0},
                    "ferp": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                # 02-01-002-09 и т.п.
                                "code": {
                                    "type": "string",
                                    "pattern": r"^\d{2}-\d{2}-\d{3}-\d{2}$"
                                },
                                "title": {"type": "string", "minLength": 3}
                            },
                            "required": ["code", "title"]
                        }
                    },
                    # Короткая сводка действий на основе выбранных ФЕРп
                    "work": {"type": "string", "maxLength": 300},
                    # Свободные атрибуты (цвет RAL, U размер, способ крепления и пр.)
                    "attributes": {"type": "object", "additionalProperties": True}
                },
                "required": ["equipment_type", "name", "quantity", "priority"]
            }
        },

        "summaries": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                # Историческое поле — не убираем
                "total_switch_ports": {"type": "integer", "minimum": 0},
                # Новое: прозрачная формула подсчёта каналов для ФЕРп 02-01-002
                "channel_components": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "switch_data_ports": {"type": "integer", "minimum": 0},   # Σ(port_count × qty) по switch/router
                        "logical_mgmt_ports": {"type": "integer", "minimum": 0}, # +1 mgmt на каждый switch/router (qty)
                        "voip_phones": {"type": "integer", "minimum": 0},
                        "ups_units": {"type": "integer", "minimum": 0},
                        "other_channels": {"type": "integer", "minimum": 0, "default": 0}
                    },
                    "required": ["switch_data_ports", "logical_mgmt_ports", "voip_phones", "ups_units"]
                },
                "total_channels": {"type": "integer", "minimum": 0},  # Итого для 02-01-002-09/-10
                "ferp_02_01_002_summary": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "base_applied": {"type": "boolean", "default": True},   # 02-01-002-09
                        "base_limit": {"type": "integer", "const": 80},         # лимит по базе
                        "extra_channels": {"type": "integer", "minimum": 0},    # 02-01-002-10
                        "notes": {"type": "string"}
                    },
                    "required": ["base_applied", "base_limit", "extra_channels"]
                },
                "notes": {"type": "string"}
            },
            "required": ["total_switch_ports"]
        },

        "pnr": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "main_tasks_5": {
                    "type": "array",
                    "minItems": 5,
                    "maxItems": 5,
                    "items": {"type": "string"}
                },
                "checklist_10_15": {
                    "type": "array",
                    "minItems": 10,
                    "maxItems": 15,
                    "items": {"type": "string"}
                }
            },
            "required": ["main_tasks_5", "checklist_10_15"]
        }
    },
    "required": ["project", "equipment", "summaries", "pnr"]
}
