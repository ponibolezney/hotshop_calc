AI_EXTRACTION_JSON_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "constants": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "k1": {"type": "number"},
                "k2": {"type": "number"},
                "k_empirical": {"type": "number"},
                "z_m": {"type": "number"},
                "ko": {"type": "number"},
                "a": {"type": "number"}
            },
            "required": ["k1", "k2", "k_empirical", "z_m", "ko", "a"]
        },
        "rooms": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "room_name": {"type": "string"},
                    "room_category": {"type": "string"},
                    "source_text": {"type": "string"},
                    "confidence": {"type": "number"},
                    "equipment": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "name": {"type": "string"},
                                "equipment_type_id": {"type": "string"},
                                "quantity": {"type": "integer"},
                                "qy_kw": {"type": "number"},
                                "ka_w_per_kw": {"type": "number"},
                                "kz": {"type": "number"},
                                "position": {"type": "string"},
                                "width_mm": {"type": "number"},
                                "depth_mm": {"type": "number"},
                                "room_name": {"type": "string"},

                                "source_text": {"type": "string"},
                                "confidence": {"type": "number"},
                                "needs_review": {"type": "boolean"}
                            },
                            "required": [
                                "name",
                                "equipment_type_id",
                                "quantity",
                                "qy_kw",
                                "ka_w_per_kw",
                                "kz",
                                "position",
                                "width_mm",
                                "depth_mm",
                                "room_name",
                                "source_text",
                                "confidence",
                                "needs_review"
                            ]
                        }
                    }
                },
                "required": [
                    "room_name",
                    "room_category",
                    "source_text",
                    "confidence",
                    "equipment"
                ]
            }
        }
    },
    "required": ["constants", "rooms"]
}