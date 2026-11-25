"""Held-out rounds of a shape's ThreeDFormat (Shape.ThreeD): every member set and read, typed and late, on an AutoShape, an
oval, a text box and a line, and what each change writes to the shape's spPr (scene3d and sp3d). The model's answers are
written down first, each case run alone, so a refusal answers that case only; a read the model refuses answers "refused" in its
own place.

Each case makes a workbook, adds the case's shape as sh, runs its statement on t, its ThreeD, and reads READS, each as its
TypeName and value or "e" and the error number, then saves the workbook and records the shape's spPr.

    python scripts/measure_three_d_round.py predict [round]
    python scripts/measure_three_d_round.py measure [round]

writes tests/fixtures/three_d_round_<round>.json, which tests/test_excel_three_d_round.py replays.
"""
from __future__ import annotations

import hashlib
import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
SEPARATOR = '%%'
#: Between what a case sets, before its book is saved and opened again, and what it does after.
AGAIN = ' || '
SHAPES = {'rect': 'ws.Shapes.AddShape(msoShapeRectangle, 20, 20, 120, 60)', 'oval': 'ws.Shapes.AddShape(msoShapeOval, 20, 20, 120, 60)',
          'text': 'ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 20, 20, 120, 60)', 'line': 'ws.Shapes.AddLine(20, 20, 140, 80)'}
READS = ['TypeName(t)', 'TypeName(t.Parent)', 't.Visible', 't.BevelTopType', 't.BevelTopDepth', 't.BevelTopInset', 't.BevelBottomType',
         't.BevelBottomDepth', 't.BevelBottomInset', 't.Depth', 't.ContourWidth', 't.ContourColor.RGB', 't.ExtrusionColor.RGB',
         't.ExtrusionColorType', 't.FieldOfView', 't.LightAngle', 't.Perspective', 't.PresetCamera', 't.PresetExtrusionDirection',
         't.PresetLighting', 't.PresetLightingDirection', 't.PresetLightingSoftness', 't.PresetMaterial', 't.PresetThreeDFormat',
         't.ProjectText', 't.RotationX', 't.RotationY', 't.RotationZ', 't.Z', 'TypeName(t.ExtrusionColor)']
ROUNDS: dict[int, list[str]] = {
    # Round 1: each scalar member set alone, some out of range, and the rotation methods.
    1: ['', 't.Visible = msoTrue', 't.Visible = msoFalse', 't.BevelTopType = msoBevelCircle', 't.BevelTopDepth = 5', 't.BevelTopInset = 4',
        't.BevelBottomType = msoBevelAngle', 't.BevelBottomDepth = 3', 't.BevelBottomInset = 2', 't.Depth = 10', 't.ContourWidth = 2',
        't.ContourColor.RGB = RGB(255, 0, 0)', 't.ExtrusionColor.RGB = RGB(0, 0, 255)', 't.ExtrusionColorType = msoExtrusionColorCustom',
        't.FieldOfView = 30', 't.LightAngle = 20', 't.Perspective = msoTrue', 't.ProjectText = msoTrue', 't.RotationX = 30',
        't.RotationY = 20', 't.RotationZ = 10', 't.Z = 5', 't.IncrementRotationX 10', 't.IncrementRotationY 10', 't.IncrementRotationZ 10',
        't.IncrementRotationHorizontal 10', 't.IncrementRotationVertical 10', 't.RotationX = 30: t.ResetRotation', 't.Depth = -1',
        't.BevelTopDepth = 2000', 't.RotationX = 400', 't.FieldOfView = 200', 't.Depth = 10: t.Visible = msoFalse',
        't.BevelTopType = msoBevelCircle: t.BevelTopDepth = 5: t.BevelTopInset = 4'],
    # Round 2: every preset table from an untouched shape, the other bevels, the bounds, and what round 1 left open.
    2: [*(f't.SetThreeDFormat {n}' for n in range(1, 21)), *(f't.SetPresetCamera {n}' for n in range(1, 63)),
        *(f't.PresetLighting = {n}' for n in range(1, 28)), *(f't.PresetLightingDirection = {n}' for n in range(1, 10)),
        *(f't.PresetLightingSoftness = {n}' for n in range(1, 4)), *(f't.PresetMaterial = {n}' for n in range(1, 16)),
        *(f't.SetExtrusionDirection {n}' for n in range(1, 10)),
        *(f't.BevelTopType = {n}' for n in (1, 2, 4, 5, 7, 8, 9, 10, 11, 12, 13)), 't.BevelTopType = 3: t.BevelTopType = 1',
        't.BevelBottomType = 12: t.BevelBottomDepth = 0', 't.Depth = -5', 't.Depth = 0', 't.Depth = 1584', 't.Depth = 1585',
        't.BevelTopDepth = 1584', 't.BevelTopDepth = 1585', 't.BevelTopDepth = -1', 't.BevelTopInset = 0', 't.FieldOfView = 180',
        't.FieldOfView = 181', 't.FieldOfView = 0', 't.LightAngle = 400', 't.LightAngle = -20', 't.Z = -5', 't.Z = 2000',
        't.ContourWidth = 2000', 't.ContourWidth = 0', 't.Perspective = msoTrue: t.FieldOfView = 30',
        't.Perspective = msoTrue: t.RotationX = 30', 't.RotationX = 30: t.Perspective = msoTrue',
        't.Perspective = msoTrue: t.Perspective = msoFalse', 't.Perspective = msoFalse', 't.ProjectText = msoFalse',
        't.Depth = 10: t.ProjectText = msoFalse', 't.ExtrusionColor.RGB = RGB(0, 0, 255): t.ExtrusionColorType = msoExtrusionColorAutomatic',
        't.Depth = 10: t.Visible = msoFalse: t.Visible = msoTrue', 't.RotationX = 30: t.Visible = msoFalse',
        't.Depth = 10: t.Visible = msoFalse: t.Depth = 20', 't.ResetRotation', 't.IncrementRotationVertical 20',
        't.IncrementRotationVertical 45', 't.RotationY = 30: t.IncrementRotationVertical 10', 't.IncrementRotationHorizontal 45',
        't.RotationX = 180', 't.RotationX = -180', 't.RotationY = 90', 't.RotationY = 270', 't.RotationZ = 190',
        't.ContourColor.ObjectThemeColor = msoThemeColorAccent2',
        't.Depth = 10: t.Z = 5: t.ContourWidth = 2: t.BevelTopType = 6: t.BevelBottomType = 4: t.ExtrusionColor.RGB = RGB(1, 2, 3): '
        't.ContourColor.RGB = RGB(4, 5, 6)', 't.RotationX = 10: t.RotationY = 20: t.RotationZ = 30: t.LightAngle = 40'],
    # Round 3: what was set, saved, the book opened again, read and changed.
    3: [f'{before}{AGAIN}{after}' for before in (
        '', 't.Depth = 10', 't.BevelTopType = msoBevelAngle: t.BevelTopDepth = 3', 't.SetThreeDFormat 5', 't.SetPresetCamera 25',
        't.PresetLighting = 20', 't.PresetLightingDirection = 3',
        't.ExtrusionColor.RGB = RGB(1, 2, 3): t.ContourColor.ObjectThemeColor = msoThemeColorAccent2', 't.RotationX = 30: t.LightAngle = 20',
        't.Perspective = msoTrue: t.FieldOfView = 30', 't.Depth = 10: t.Visible = msoFalse', 't.Depth = 10: t.Visible = msoFalse: t.Depth = 20',
        't.BevelTopType = msoBevelNone', 't.ProjectText = msoFalse')
        for after in ('', 't.Depth = 20', 't.RotationY = 15', 't.Visible = msoFalse', 't.Visible = msoTrue')],
    # Round 4, held out: presets over what is set, hiding and showing in turn, the bounds, a ShapeRange, and where ProjectText goes.
    4: ['t.ProjectText = msoFalse', 't.Depth = 5: t.ProjectText = msoFalse: t.ProjectText = msoTrue', 't.SetPresetCamera 25: t.RotationX = 10',
        't.SetPresetCamera 49: t.RotationZ = 5', 't.SetPresetCamera 3: t.SetPresetCamera 40', 't.RotationX = 20: t.SetPresetCamera 19',
        't.SetThreeDFormat 7: t.Depth = 5', 't.SetThreeDFormat 2: t.SetThreeDFormat 9', 't.Depth = 5: t.SetThreeDFormat 1',
        't.PresetLighting = 3: t.PresetLightingDirection = 6', 't.PresetLighting = 7: t.PresetLightingSoftness = 1',
        't.PresetLightingDirection = 3: t.PresetLighting = 15', 't.PresetLighting = 14: t.PresetLightingDirection = 2',
        't.Depth = 5: t.PresetMaterial = 3', 't.PresetMaterial = 6: t.PresetMaterial = 8', 't.RotationX = 10: t.SetExtrusionDirection 4',
        't.Depth = 5: t.SetExtrusionDirection 2', 't.SetPresetCamera 12: t.SetExtrusionDirection 9', 't.SetPresetCamera 4: t.Perspective = msoTrue',
        't.SetPresetCamera 16: t.Perspective = msoFalse', 't.SetPresetCamera 48: t.FieldOfView = 50', 't.SetPresetCamera 11: t.FieldOfView = 50',
        't.Perspective = msoTrue: t.ResetRotation', 't.Depth = 5: t.Visible = msoFalse: t.RotationX = 10: t.Visible = msoTrue',
        't.RotationX = 10: t.Visible = msoFalse: t.Depth = 5: t.Visible = msoTrue', 't.Depth = 5: t.RotationX = 10: t.Visible = msoFalse: t.Visible = msoTrue',
        't.Depth = 5: t.Visible = msoFalse: t.Depth = 9: t.Visible = msoFalse', 't.ExtrusionColor.ObjectThemeColor = msoThemeColorAccent4',
        't.ContourColor.RGB = RGB(9, 8, 7): t.ContourWidth = 0', 't.ExtrusionColor.RGB = RGB(9, 8, 7): t.Depth = 0', 't.BevelBottomType = msoBevelNone',
        't.BevelTopInset = 3: t.BevelTopType = msoBevelNone: t.BevelTopDepth = 2', 't.Z = 3: t.Visible = msoFalse',
        't.LightAngle = 45: t.PresetLighting = 20', 't.PresetLighting = 20: t.LightAngle = 45', 't.RotationY = 350: t.IncrementRotationY 20',
        't.IncrementRotationX -30', 't.IncrementRotationZ 400', 't.RotationY = 40: t.IncrementRotationVertical 40', 't.Depth = 1585.5',
        't.ContourWidth = 1584', 't.FieldOfView = 179', 't.Perspective = msoTrue: t.FieldOfView = 179.5', 't.Z = -2000',
        'Set t = ws.Shapes.Range(Array("S")).ThreeD: t.Depth = 7', 'Set t = ws.Shapes.Range(Array("S")).ThreeD: t.SetPresetCamera 20',
        'Set t = ws.Shapes.Range(Array("S")).ThreeD',
        't.BevelTopType = msoBevelCircle: t.BevelBottomType = msoBevelCross: t.BevelTopDepth = 0: t.BevelBottomInset = 1584',
        't.SetThreeDFormat 12: t.Visible = msoFalse', 't.SetPresetCamera 30: t.RotationY = 5', 't.PresetLightingSoftness = 2: t.PresetLightingDirection = 8',
        't.SetExtrusionDirection 6: t.PresetLightingDirection = 1', 't.PresetMaterial = 11: t.BevelTopType = msoBevelNone',
        't.Depth = 3: t.ContourWidth = 1: t.ExtrusionColor.RGB = RGB(1, 1, 1): t.ExtrusionColorType = msoExtrusionColorAutomatic',
        't.SetPresetCamera 59', 't.SetThreeDFormat 20', 't.SetPresetCamera 14: t.RotationX = 15'],
    # Round 5, held out on the code that answered round 4: the same ground with other values and orders, and files opened again.
    5: ['t.BevelBottomType = msoBevelDivot: t.BevelBottomDepth = 4.5: t.BevelBottomInset = 2.25', 't.BevelTopType = msoBevelArtDeco: t.BevelTopType = msoBevelSlope',
        't.Depth = 3: t.BevelTopType = msoBevelRiblet: t.SetThreeDFormat 11', 't.ContourWidth = 3: t.SetThreeDFormat 4', 't.RotationZ = 45: t.SetThreeDFormat 6',
        't.RotationX = 15: t.PresetMaterial = 9', 't.BevelTopType = msoBevelNone: t.PresetLightingSoftness = 2', 't.LightAngle = 90',
        't.LightAngle = 180: t.PresetLighting = 25', 't.LightAngle = 270', 't.LightAngle = 315: t.PresetLighting = 18', 't.LightAngle = 30: t.PresetLighting = 16',
        't.PresetLightingDirection = 2: t.PresetLighting = 17', 't.PresetLightingDirection = 8: t.PresetLighting = 21',
        't.PresetLighting = 9: t.LightAngle = 45', 't.PresetLighting = 6: t.PresetLightingDirection = 9', 't.PresetLighting = 12: t.PresetLightingSoftness = 3',
        't.SetPresetCamera 8: t.Perspective = msoTrue', 't.SetPresetCamera 18: t.Perspective = msoFalse', 't.SetPresetCamera 5: t.Perspective = msoTrue: t.FieldOfView = 40',
        't.RotationY = 25: t.ResetRotation', 't.RotationZ = 15: t.Perspective = msoTrue: t.ResetRotation', 't.Depth = 8: t.SetExtrusionDirection 7',
        't.Depth = 8: t.SetPresetCamera 13: t.SetExtrusionDirection 1', 't.SetPresetCamera 26: t.RotationZ = 30', 't.SetPresetCamera 55: t.RotationX = 5',
        't.SetPresetCamera 21: t.IncrementRotationY 5', 't.SetPresetCamera 61', 't.SetPresetCamera 62: t.FieldOfView = 20', 't.SetPresetCamera 44: t.FieldOfView = 20',
        't.Z = 12: t.Visible = msoFalse: t.Visible = msoTrue', 't.LightAngle = 60: t.Visible = msoFalse: t.Z = 2: t.Visible = msoTrue',
        't.ContourWidth = 4: t.Visible = msoFalse: t.Visible = msoFalse', 't.SetThreeDFormat 3: t.Visible = msoFalse: t.Depth = 7: t.Visible = msoTrue',
        't.ExtrusionColor.ObjectThemeColor = msoThemeColorAccent6: t.ContourColor.ObjectThemeColor = msoThemeColorAccent1',
        't.ExtrusionColor.RGB = RGB(10, 20, 30): t.ExtrusionColorType = msoExtrusionColorCustom', 't.Depth = 0: t.ContourWidth = 0',
        't.ProjectText = msoFalse: t.Depth = 4', 't.ProjectText = msoFalse: t.Visible = msoFalse', 't.IncrementRotationHorizontal -15: t.IncrementRotationZ -20',
        't.RotationY = 100: t.IncrementRotationVertical 30', 't.FieldOfView = 50', 't.Perspective = msoTrue: t.FieldOfView = 50.5', 't.Depth = -1585.5',
        't.BevelBottomDepth = 1584: t.BevelBottomInset = 1584', 't.Z = 2000: t.Z = -2000', 't.ContourWidth = 1585',
        'Set t = ws.Shapes.Range(Array("S")).ThreeD: t.ExtrusionColor.RGB = RGB(4, 3, 2): t.BevelTopType = msoBevelConvex',
        'Set t = ws.Shapes.Range(Array("S")).ThreeD: t.PresetLighting = 22: t.IncrementRotationX 12',
        f't.SetThreeDFormat 13{AGAIN}t.Depth = 3', f't.PresetLighting = 4{AGAIN}t.PresetLightingSoftness = 1', f't.SetPresetCamera 50{AGAIN}t.RotationZ = 10',
        f't.ProjectText = msoFalse{AGAIN}t.ProjectText = msoTrue', f't.ProjectText = msoFalse: t.Depth = 3{AGAIN}', f't.Z = 4: t.Visible = msoFalse{AGAIN}t.Visible = msoTrue',
        f't.PresetMaterial = 12{AGAIN}t.PresetMaterial = 2', f't.Perspective = msoTrue: t.FieldOfView = 35{AGAIN}t.ResetRotation',
        f't.SetExtrusionDirection 3{AGAIN}t.PresetLightingDirection = 7', f't.LightAngle = 45{AGAIN}'],
    # Round 6, held out on the code that answered round 5.
    6: ['t.BevelTopType = msoBevelHardEdge: t.BevelTopInset = 7: t.BevelBottomType = msoBevelSoftRound', 't.BevelBottomDepth = 2: t.BevelBottomType = msoBevelNone',
        't.Z = 6: t.SetThreeDFormat 14', 't.ExtrusionColor.RGB = RGB(5, 5, 5): t.SetThreeDFormat 8', 't.SetThreeDFormat 17: t.ContourWidth = 2',
        't.SetThreeDFormat 10: t.RotationZ = 20', 't.SetThreeDFormat 19: t.PresetLighting = 14', 't.SetThreeDFormat 1: t.BevelTopType = msoBevelAngle',
        't.LightAngle = 135', 't.LightAngle = 225: t.PresetLighting = 19', 't.LightAngle = 359', 't.LightAngle = 90: t.PresetLighting = 5',
        't.PresetLighting = 11: t.LightAngle = 270', 't.PresetLighting = 26: t.LightAngle = 315', 't.PresetLighting = 2: t.PresetLightingSoftness = 2',
        't.PresetLighting = 10: t.PresetLightingDirection = 4', 't.PresetLightingSoftness = 1: t.PresetLightingDirection = 5',
        't.PresetLightingDirection = 7: t.PresetLightingSoftness = 2', 't.Depth = 2: t.PresetLightingDirection = 6', 't.Z = 1: t.PresetLightingSoftness = 3',
        't.RotationY = 10: t.PresetMaterial = 13', 't.PresetMaterial = 4: t.Depth = 6', 't.BevelTopType = msoBevelNone: t.PresetMaterial = 15',
        't.SetPresetCamera 2: t.Perspective = msoTrue: t.FieldOfView = 25', 't.SetPresetCamera 13: t.Perspective = msoFalse: t.RotationY = 5',
        't.SetPresetCamera 19: t.Perspective = msoTrue', 't.SetPresetCamera 10: t.ResetRotation', 't.RotationX = 5: t.RotationY = 5: t.ResetRotation',
        't.SetPresetCamera 15: t.RotationY = 12', 't.SetPresetCamera 24: t.IncrementRotationZ 15', 't.SetPresetCamera 53: t.RotationX = 0',
        't.SetPresetCamera 60: t.FieldOfView = 10', 't.SetPresetCamera 7: t.FieldOfView = 10', 't.SetPresetCamera 35: t.SetPresetCamera 48',
        't.Depth = 4: t.SetExtrusionDirection 5', 't.SetExtrusionDirection 8: t.SetExtrusionDirection 3', 't.Depth = 4: t.SetPresetCamera 17: t.SetExtrusionDirection 6',
        't.RotationY = 20: t.Visible = msoFalse: t.Visible = msoTrue', 't.Depth = 6: t.Visible = msoFalse: t.Z = 2: t.Visible = msoFalse: t.Visible = msoTrue',
        't.SetPresetCamera 30: t.Visible = msoFalse', 't.Visible = msoFalse: t.Visible = msoTrue', 't.BevelTopType = msoBevelNone: t.Visible = msoFalse',
        't.ContourColor.ObjectThemeColor = msoThemeColorAccent5: t.ContourWidth = 3', 't.ExtrusionColor.RGB = RGB(200, 100, 50): t.Depth = 9',
        't.ProjectText = msoFalse: t.ProjectText = msoFalse', 't.ProjectText = msoTrue: t.Visible = msoFalse', 't.RotationX = -30: t.IncrementRotationHorizontal 10',
        't.RotationY = 300: t.IncrementRotationVertical -20', 't.IncrementRotationY -5', 't.RotationZ = -10: t.IncrementRotationZ 25', 't.FieldOfView = 49.5',
        't.Perspective = msoTrue: t.FieldOfView = 0', 't.Depth = 0.5', 't.BevelTopDepth = 0.25: t.BevelTopInset = 1583.5', 't.Z = 1999.5', 't.ContourWidth = 0.75',
        'Set t = ws.Shapes.Range(Array("S")).ThreeD: t.SetThreeDFormat 3: t.ContourColor.RGB = RGB(1, 2, 3)',
        'Set t = ws.Shapes.Range(Array("S")).ThreeD: t.Visible = msoFalse',
        f't.Depth = 6: t.BevelTopType = msoBevelCoolSlant{AGAIN}t.BevelTopDepth = 2', f't.SetPresetCamera 56{AGAIN}t.RotationY = 0',
        f't.PresetLighting = 23: t.LightAngle = 90{AGAIN}', f't.ExtrusionColor.ObjectThemeColor = msoThemeColorAccent3{AGAIN}t.ExtrusionColorType = msoExtrusionColorAutomatic',
        f't.SetThreeDFormat 16{AGAIN}t.SetExtrusionDirection 4', f't.RotationX = 25: t.Visible = msoFalse{AGAIN}t.Depth = 5',
        f't.ProjectText = msoFalse{AGAIN}t.Visible = msoFalse', f't.Perspective = msoTrue{AGAIN}t.Perspective = msoFalse'],
    # Round 7, held out on the code that answered round 6.
    7: ['t.BevelTopType = msoBevelCross: t.BevelBottomType = msoBevelArtDeco: t.BevelBottomInset = 9', 't.BevelTopDepth = 3.5: t.BevelTopType = msoBevelNone',
        't.ContourWidth = 5: t.SetThreeDFormat 18', 't.SetThreeDFormat 15: t.PresetMaterial = 7', 't.SetThreeDFormat 6: t.LightAngle = 30',
        't.SetThreeDFormat 12: t.RotationY = 10', 't.SetThreeDFormat 9: t.Depth = 36', 't.Z = 8: t.SetThreeDFormat 2',
        't.LightAngle = 100', 't.LightAngle = 45: t.PresetLighting = 27', 't.LightAngle = 315', 't.PresetLighting = 24: t.LightAngle = 135',
        't.PresetLighting = 1: t.LightAngle = 90', 't.PresetLighting = 4: t.LightAngle = 315', 't.PresetLighting = 8: t.PresetLightingSoftness = 1',
        't.PresetLighting = 5: t.PresetLightingDirection = 3', 't.PresetLighting = 3: t.PresetLightingDirection = 8', 't.PresetLighting = 12: t.PresetLightingDirection = 1',
        't.PresetLightingDirection = 9: t.PresetLightingSoftness = 1', 't.PresetLightingSoftness = 3: t.PresetLightingDirection = 2',
        't.PresetLighting = 16: t.PresetLighting = 2', 't.ContourWidth = 2: t.PresetMaterial = 10', 't.PresetMaterial = 14: t.SetExtrusionDirection 2',
        't.SetPresetCamera 6: t.Perspective = msoTrue: t.RotationX = 5', 't.SetPresetCamera 17: t.Perspective = msoFalse: t.Perspective = msoTrue',
        't.SetPresetCamera 14: t.FieldOfView = 15', 't.SetPresetCamera 52: t.FieldOfView = 35', 't.SetPresetCamera 41: t.FieldOfView = 35',
        't.RotationZ = 20: t.ResetRotation', 't.SetPresetCamera 14: t.RotationY = 10: t.ResetRotation', 't.SetPresetCamera 29: t.RotationY = 8',
        't.SetPresetCamera 58: t.RotationZ = 1', 't.SetPresetCamera 22: t.IncrementRotationX 3', 't.SetPresetCamera 45: t.SetPresetCamera 19',
        't.SetPresetCamera 9: t.SetExtrusionDirection 7', 't.ContourWidth = 1: t.SetExtrusionDirection 9', 't.Depth = 5: t.SetPresetCamera 11: t.SetExtrusionDirection 8',
        't.LightAngle = 20: t.Visible = msoFalse: t.Visible = msoTrue', 't.Depth = 3: t.RotationX = 5: t.Visible = msoFalse: t.Depth = 6: t.Visible = msoTrue',
        't.SetThreeDFormat 7: t.Visible = msoFalse: t.Visible = msoTrue', 't.ContourColor.RGB = RGB(1, 1, 1): t.Visible = msoFalse',
        't.ExtrusionColor.ObjectThemeColor = msoThemeColorAccent2: t.ExtrusionColor.RGB = RGB(7, 7, 7)', 't.ContourColor.ObjectThemeColor = msoThemeColorAccent6: t.ContourWidth = 0',
        't.ExtrusionColorType = msoExtrusionColorAutomatic', 't.ProjectText = msoFalse: t.Depth = 0', 't.Depth = 2: t.ProjectText = msoFalse: t.Visible = msoFalse: t.Visible = msoTrue',
        't.RotationX = 60: t.IncrementRotationHorizontal 20', 't.IncrementRotationX 370', 't.RotationY = -20: t.IncrementRotationY -15', 't.IncrementRotationZ -5: t.IncrementRotationZ 5',
        't.FieldOfView = 25', 't.Perspective = msoTrue: t.FieldOfView = 49', 't.Depth = 1500', 't.BevelBottomDepth = 0.5: t.BevelBottomInset = 0', 't.Z = -1999', 't.ContourWidth = 1000',
        'Set t = ws.Shapes.Range(Array("S")).ThreeD: t.PresetMaterial = 5: t.LightAngle = 10',
        'Set t = ws.Shapes.Range(Array("S")).ThreeD: t.Perspective = msoTrue: t.ContourColor.ObjectThemeColor = msoThemeColorAccent3',
        f't.BevelBottomType = msoBevelConvex: t.BevelBottomDepth = 3{AGAIN}t.BevelBottomType = msoBevelNone', f't.SetPresetCamera 33{AGAIN}',
        f't.PresetLighting = 10: t.PresetLightingDirection = 9{AGAIN}t.PresetLightingSoftness = 3', f't.ContourColor.ObjectThemeColor = msoThemeColorAccent4{AGAIN}t.ContourWidth = 4',
        f't.SetThreeDFormat 20{AGAIN}t.PresetMaterial = 3', f't.Z = 7: t.RotationY = 5: t.Visible = msoFalse{AGAIN}t.Z = 9',
        f't.ProjectText = msoFalse: t.Visible = msoFalse{AGAIN}t.Visible = msoTrue', f't.SetPresetCamera 15: t.FieldOfView = 30{AGAIN}t.Perspective = msoFalse'],
    # Round 8, held out on the code that answered round 7.
    8: ['t.BevelTopType = msoBevelDivot: t.BevelTopDepth = 11: t.BevelBottomType = msoBevelRelaxedInset: t.BevelBottomDepth = 1',
        't.BevelBottomInset = 6: t.BevelBottomType = msoBevelNone: t.BevelTopType = msoBevelNone', 't.Depth = 12: t.SetThreeDFormat 16', 't.SetThreeDFormat 13: t.ContourWidth = 6',
        't.SetThreeDFormat 5: t.PresetLighting = 26', 't.SetThreeDFormat 11: t.Z = 2', 't.SetThreeDFormat 3: t.Depth = 36', 't.ContourColor.RGB = RGB(3, 3, 3): t.SetThreeDFormat 10',
        't.LightAngle = 0', 't.LightAngle = 45', 't.LightAngle = 270: t.PresetLighting = 23', 't.PresetLighting = 21: t.LightAngle = 90', 't.PresetLighting = 9: t.LightAngle = 315',
        't.PresetLighting = 11: t.LightAngle = 45', 't.PresetLighting = 12: t.LightAngle = 90', 't.PresetLighting = 7: t.LightAngle = 20',
        't.PresetLighting = 6: t.PresetLightingSoftness = 3', 't.PresetLighting = 10: t.PresetLightingSoftness = 2', 't.PresetLighting = 1: t.PresetLightingDirection = 5',
        't.PresetLighting = 7: t.PresetLightingDirection = 2', 't.PresetLighting = 11: t.PresetLightingDirection = 6', 't.PresetLightingDirection = 4: t.PresetLightingSoftness = 3',
        't.PresetLightingSoftness = 2: t.PresetLightingSoftness = 1', 't.PresetLighting = 17: t.PresetLighting = 9', 't.Z = 3: t.PresetMaterial = 5', 't.PresetMaterial = 1: t.SetExtrusionDirection 8',
        't.SetPresetCamera 1: t.Perspective = msoTrue: t.RotationZ = 7', 't.SetPresetCamera 12: t.Perspective = msoFalse: t.FieldOfView = 20',
        't.SetPresetCamera 10: t.FieldOfView = 45', 't.SetPresetCamera 57: t.FieldOfView = 12', 't.SetPresetCamera 47: t.FieldOfView = 12', 't.RotationX = 7: t.RotationZ = 3: t.ResetRotation',
        't.Perspective = msoTrue: t.RotationY = 9: t.ResetRotation', 't.SetPresetCamera 34: t.RotationZ = 6', 't.SetPresetCamera 54: t.RotationY = 2', 't.SetPresetCamera 23: t.IncrementRotationZ 4',
        't.SetPresetCamera 42: t.SetPresetCamera 14', 't.SetPresetCamera 3: t.SetExtrusionDirection 5', 't.Z = 2: t.SetExtrusionDirection 1', 't.Depth = 5: t.SetPresetCamera 10: t.SetExtrusionDirection 4',
        't.Perspective = msoTrue: t.Visible = msoFalse: t.Visible = msoTrue', 't.Depth = 3: t.Visible = msoFalse: t.BevelTopType = msoBevelSlope: t.Visible = msoTrue',
        't.SetThreeDFormat 18: t.Visible = msoFalse', 't.ExtrusionColor.RGB = RGB(2, 2, 2): t.Visible = msoFalse: t.Visible = msoTrue',
        't.ContourColor.ObjectThemeColor = msoThemeColorAccent3: t.ContourColor.RGB = RGB(6, 6, 6)', 't.ExtrusionColor.ObjectThemeColor = msoThemeColorAccent5: t.Depth = 0',
        't.Depth = 4: t.ExtrusionColorType = msoExtrusionColorAutomatic', 't.ProjectText = msoFalse: t.BevelTopType = msoBevelAngle',
        't.ProjectText = msoTrue: t.ProjectText = msoFalse: t.Depth = 3', 't.RotationX = 33: t.IncrementRotationX -50', 't.IncrementRotationY 365', 't.RotationZ = 359: t.IncrementRotationZ 2',
        't.RotationY = 45: t.RotationX = -45: t.RotationZ = 90', 't.FieldOfView = 10', 't.Perspective = msoTrue: t.FieldOfView = 10', 't.Depth = 777', 't.BevelTopInset = 0.5: t.BevelTopDepth = 0',
        't.Z = 1', 't.ContourWidth = 100',
        'Set t = ws.Shapes.Range(Array("S")).ThreeD: t.SetThreeDFormat 14: t.ExtrusionColor.ObjectThemeColor = msoThemeColorAccent1',
        'Set t = ws.Shapes.Range(Array("S")).ThreeD: t.BevelBottomType = msoBevelCircle: t.RotationY = 7',
        f't.BevelTopType = msoBevelSoftRound: t.BevelTopInset = 2{AGAIN}t.BevelTopType = msoBevelCircle', f't.SetPresetCamera 59{AGAIN}t.SetPresetCamera 19',
        f't.PresetLighting = 8: t.PresetLightingDirection = 7{AGAIN}t.PresetLightingSoftness = 2', f't.ExtrusionColor.RGB = RGB(9, 9, 9){AGAIN}t.ExtrusionColorType = msoExtrusionColorAutomatic',
        f't.SetThreeDFormat 7{AGAIN}t.Visible = msoFalse', f't.LightAngle = 270: t.Visible = msoFalse{AGAIN}t.LightAngle = 90',
        f't.ProjectText = msoFalse: t.Depth = 2{AGAIN}t.Visible = msoFalse', f't.SetPresetCamera 16: t.FieldOfView = 40{AGAIN}t.FieldOfView = 20'],
    # Round 9, held out on the code that answered round 8.
    9: ['t.BevelTopType = msoBevelRiblet: t.BevelTopInset = 4: t.BevelBottomType = msoBevelSlope', 't.BevelBottomType = msoBevelCoolSlant: t.BevelBottomInset = 12',
        't.Depth = 8: t.SetThreeDFormat 4', 't.SetThreeDFormat 20: t.ContourWidth = 3', 't.SetThreeDFormat 2: t.PresetLighting = 9', 't.SetThreeDFormat 19: t.RotationX = 15',
        't.SetThreeDFormat 6: t.Z = 5', 't.ExtrusionColor.ObjectThemeColor = msoThemeColorAccent4: t.SetThreeDFormat 12',
        't.LightAngle = 90', 't.LightAngle = 180', 't.LightAngle = 300', 't.PresetLighting = 22: t.LightAngle = 45', 't.PresetLighting = 13: t.LightAngle = 270',
        't.PresetLighting = 25: t.LightAngle = 90', 't.PresetLighting = 4: t.LightAngle = 45', 't.PresetLighting = 1: t.LightAngle = 270',
        't.PresetLighting = 15: t.PresetLightingSoftness = 1', 't.PresetLighting = 3: t.PresetLightingSoftness = 3', 't.PresetLighting = 4: t.PresetLightingDirection = 3',
        't.PresetLighting = 2: t.PresetLightingDirection = 9', 't.PresetLighting = 13: t.PresetLightingDirection = 4', 't.PresetLightingDirection = 1: t.PresetLightingSoftness = 1',
        't.PresetLightingSoftness = 3: t.PresetLightingSoftness = 2', 't.PresetLighting = 20: t.PresetLighting = 6', 't.Depth = 7: t.PresetMaterial = 9', 't.PresetMaterial = 6: t.SetExtrusionDirection 4',
        't.SetPresetCamera 4: t.Perspective = msoTrue: t.RotationY = 6', 't.SetPresetCamera 20: t.Perspective = msoFalse: t.RotationX = 8',
        't.SetPresetCamera 25: t.FieldOfView = 30', 't.SetPresetCamera 49: t.FieldOfView = 20', 't.SetPresetCamera 38: t.FieldOfView = 20', 't.RotationY = 15: t.RotationZ = 4: t.ResetRotation',
        't.SetPresetCamera 27: t.RotationX = 4: t.ResetRotation', 't.SetPresetCamera 44: t.RotationZ = 9', 't.SetPresetCamera 61: t.RotationY = 3', 't.SetPresetCamera 31: t.IncrementRotationY 6',
        't.SetPresetCamera 50: t.SetPresetCamera 5', 't.SetPresetCamera 8: t.SetExtrusionDirection 2', 't.BevelTopType = msoBevelAngle: t.SetExtrusionDirection 6',
        't.Depth = 6: t.SetPresetCamera 2: t.SetExtrusionDirection 9',
        't.Z = 3: t.Visible = msoFalse: t.Visible = msoTrue', 't.Depth = 4: t.Visible = msoFalse: t.RotationY = 5: t.Visible = msoTrue',
        't.SetThreeDFormat 11: t.Visible = msoFalse', 't.ContourColor.RGB = RGB(4, 4, 4): t.ContourWidth = 2: t.Visible = msoFalse: t.Visible = msoTrue',
        't.ExtrusionColor.ObjectThemeColor = msoThemeColorAccent6: t.ExtrusionColor.ObjectThemeColor = msoThemeColorAccent2', 't.ContourColor.RGB = RGB(8, 8, 8): t.ContourWidth = 0',
        't.ExtrusionColor.RGB = RGB(1, 2, 3): t.ExtrusionColorType = msoExtrusionColorAutomatic', 't.ProjectText = msoFalse: t.SetThreeDFormat 5',
        't.ProjectText = msoFalse: t.Visible = msoFalse: t.Visible = msoTrue', 't.RotationX = 12: t.IncrementRotationX 30', 't.IncrementRotationY -10', 't.RotationZ = 5: t.IncrementRotationZ -10',
        't.RotationX = 200: t.RotationY = 100', 't.FieldOfView = 50', 't.Perspective = msoTrue: t.FieldOfView = 50', 't.Depth = 1585', 't.BevelBottomInset = 1584: t.BevelBottomDepth = 1584',
        't.Z = -3', 't.ContourWidth = 1584',
        'Set t = ws.Shapes.Range(Array("S")).ThreeD: t.SetPresetCamera 12: t.Depth = 4',
        'Set t = ws.Shapes.Range(Array("S")).ThreeD: t.BevelTopType = msoBevelArtDeco: t.PresetLighting = 7',
        f't.BevelTopType = msoBevelRiblet: t.BevelTopDepth = 5{AGAIN}t.BevelBottomType = msoBevelAngle', f't.SetPresetCamera 26{AGAIN}t.RotationX = 10',
        f't.PresetLighting = 11: t.PresetLightingSoftness = 1{AGAIN}t.LightAngle = 0', f't.ContourColor.RGB = RGB(9, 8, 7): t.ContourWidth = 2{AGAIN}t.ContourWidth = 0',
        f't.SetThreeDFormat 13{AGAIN}t.Depth = 0', f't.Depth = 5: t.Visible = msoFalse{AGAIN}t.Visible = msoTrue',
        f't.ProjectText = msoFalse: t.SetPresetCamera 3{AGAIN}t.ProjectText = msoTrue', f't.Perspective = msoTrue: t.RotationY = 4{AGAIN}t.ResetRotation'],
}
#: The shapes of each round where not all of them.
ROUND_SHAPES = {2: ['rect', 'text'], 3: ['rect', 'text']}
ROUND = 1


def out() -> Path:
    return ROOT / f'tests/fixtures/three_d_round_{ROUND}.json'


def cases() -> list[list[Any]]:
    return [[shape, statement, typed] for typed in (False, True) for shape in ROUND_SHAPES.get(ROUND, SHAPES) for statement in ROUNDS[ROUND]]


#: The reads a model run makes, where it reads one of READS alone; None for all of them.
CHOSEN: list[str] | None = None


def _case(case: list[Any], index: int) -> str:
    shape, statement, _typed = case
    lines = ['    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', f'    Set sh = {SHAPES[shape]}', '    sh.Name = "S"',
             '    Set t = sh.ThreeD']
    if AGAIN in statement:
        before, statement = statement.split(AGAIN)
        lines += [f'    {before}', f'    wb.SaveAs folder & Application.PathSeparator & "a{index}.xlsx", xlOpenXMLWorkbook',
                  '    wb.Close SaveChanges:=False', f'    Set wb = Workbooks.Open(folder & Application.PathSeparator & "a{index}.xlsx")',
                  '    Set ws = wb.Worksheets(1)', '    Set sh = ws.Shapes("S")', '    Set t = sh.ThreeD']
    lines += ['    Err.Clear', f'    {statement}', '    s = "e" & Err.Number']
    for read in READS if CHOSEN is None else CHOSEN:
        lines += ['    Err.Clear: v = Empty', f'    v = {read}', '    s = s & "|" & Show(v)']
    lines += [f'    wb.SaveAs folder & Application.PathSeparator & "t{index}.xlsx", xlOpenXMLWorkbook', '    wb.Close SaveChanges:=False',
              '    Set t = Nothing: Set sh = Nothing', f'    out = out & "{SEPARATOR}" & s']
    return '\n'.join(lines) + '\n'


HELPERS = '''Private Function Show(ByVal v As Variant) As String
    If Err.Number Then Show = "e" & Err.Number: Exit Function
    Show = TypeName(v) & ":" & CStr(v)
End Function
'''
PART = 8


def macro(chosen: list[tuple[int, list[Any]]]) -> str:
    typed = chosen[0][1][2]
    declare = ('    Dim wb As Workbook, ws As Worksheet, sh As Shape, t As ThreeDFormat, v As Variant, s As String\n' if typed else
               '    Dim wb As Object, ws As Object, sh As Object, t As Object, v As Variant, s As String\n')
    parts = [chosen[at:at + PART] for at in range(0, len(chosen), PART)]
    procedures = ''.join(f'Private Sub Part{n}(ByRef out As String, ByVal folder As String)\n{declare}    On Error Resume Next\n'
                         f'{"".join(_case(case, index) for index, case in part)}End Sub\n' for n, part in enumerate(parts))
    calls = ''.join(f'    Part{n} out, folder\n' for n in range(len(parts)))
    return (f'{HELPERS}{procedures}Function Probe(ByVal folder As String) As String\n    Dim out As String\n'
            '    Application.DisplayAlerts = False\n'
            f'{calls}    Application.DisplayAlerts = True\n    Probe = Mid(out, {len(SEPARATOR) + 1})\nEnd Function\n')


def shape_xml(folder: Path, index: int) -> str:
    """The saved shape's spPr, and from round 4 its text's bodyPr after it, or "" where the book or its drawing is not there."""
    try:
        with zipfile.ZipFile(folder / f't{index}.xlsx') as package:
            drawing = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    except (FileNotFoundError, KeyError):
        return ''
    found = re.findall(r'<xdr:spPr>.*?</xdr:spPr>|<xdr:spPr/>', drawing, re.S)
    text = re.findall(r'<a:bodyPr\b[^>]*/>|<a:bodyPr\b.*?</a:bodyPr>', drawing, re.S) if ROUND >= 4 else []
    return (found[0] if found else '') + (text[0] if text else '')


def _run(case: list[Any], index: int, folder: Path, reads: list[str] | None) -> list[str] | None:
    """The model's answer to the case with these reads, split, or None where it refuses."""
    global CHOSEN
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBACompileError, VBAUnsupportedError

    CHOSEN = reads
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    try:
        app.add_module(macro([(index, case)]), name='H')
        return str(app.run('H.Probe', str(folder))).split('|')
    except VBACompileError:
        return ['does not compile']
    except VBAUnsupportedError:
        return None
    finally:
        CHOSEN = None


def alone(case: list[Any], index: int, folder: Path) -> list[str]:
    """The model's answer to one case run by itself, and the shape's saved spPr; where a read is refused, each read alone."""
    whole = _run(case, index, folder, None)
    if whole is not None:
        return ['|'.join(whole), shape_xml(folder, index) if whole != ['does not compile'] else '']
    first = _run(case, index, folder, READS[:1])
    if first is None:
        return ['refused', '']
    xml = shape_xml(folder, index)
    reads = []
    for read in READS:
        got = _run(case, index, folder, [read])
        reads.append('refused' if got is None else got[1])
    return ['|'.join([first[0], *reads]), xml]


def predict() -> None:
    sys.path.insert(0, str(ROOT / 'src'))
    every = cases()
    with tempfile.TemporaryDirectory() as folder:
        model = [alone(case, index, Path(folder)) for index, case in enumerate(every)]
    text = json.dumps({'reads': READS, 'cases': every, 'model': model}, indent=1) + '\n'
    out().write_text(text, encoding='utf-8', newline='\n')
    print(len(every), 'cases; predictions sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())


def measure() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, Any] = json.loads(out().read_text(encoding='utf-8'))
    every: list[list[Any]] = record['cases']
    answers: list[list[str]] = []
    height = 'Function H() As String\n    H = CStr(ActiveSheet.StandardHeight)\nEnd Function\n'
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel, tempfile.TemporaryDirectory() as folder:
        excel.new_document()
        heights = [str(excel.run_vba(height, 'H', line_numbers=False).value)]
        for typed in (False, True):
            chosen = [(index, case) for index, case in enumerate(every) if case[2] is typed]
            for at in range(0, len(chosen), 24):
                result = excel.run_vba(macro(chosen[at:at + 24]), 'Probe', (folder,), timeout=600, line_numbers=False)
                assert result.ok, (at, result.error, [one.message for one in result.dialogs])
                for (index, _case_), value in zip(chosen[at:at + 24], str(result.value).split(SEPARATOR)):
                    answers.append([value, shape_xml(Path(folder), index)])
                print(typed, at + 24, flush=True)
        heights.append(str(excel.run_vba(height, 'H', line_numbers=False).value))
    record.update({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': heights, 'excel': answers})
    out().write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    wrong = sum(model != excel and model[0] != 'refused' for model, excel in zip(record['model'], answers))
    print('StandardHeight', heights, 'model wrong', wrong, 'refused', sum(model[0] == 'refused' for model in record['model']), 'of', len(answers))


if __name__ == '__main__':
    ROUND = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
