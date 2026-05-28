# -*- coding: utf-8 -*-
import re

path = r'C:\Users\rcdga\planeacion-produccion\Mod_Planificador.html'
content = open(path, encoding='utf-8').read()

# Buscar y capturar el bloque exacto con regex
pattern = (
    r"(  var _invMax = o\.invMax \|\| 0;\n"
    r"  var _invExist = o\.invExist \|\| 0;\n"
    r"  var _claseInv = '';\n"
    r"  if \(_invMax > 0\) \{[\s\S]*?"
    r"    else if \(_total < _invMax \* 0\.95\) _claseInv = ' planif-bajomax';\n"
    r"  \})"
)

m = re.search(pattern, content)
if m:
    print('C1 ENCONTRADO')
    print(repr(m.group(0)[:100]))
    old = m.group(0)
    new = """  var _invMax = o.invMax || 0;
  var _invExist = o.invExist || 0;
  var _claseInv = '';
  if (_invMax > 0) {
    var _restanKg = Math.max(0, (o.sol||0) - (o.prod||0));
    var _restanUnid = _restanKg;
    var _esVar = o.invEsVarilla === true || String(o.tipo||'').toUpperCase().indexOf('VARILLA') > -1;
    var _invBack2 = _esVar ? (parseFloat(o.invBack) || 0) : 0;
    var _invExNeg2 = _esVar ? ((o.invExistNeg !== null && o.invExistNeg !== undefined) ? (parseFloat(o.invExistNeg) || 0) : 0) : 0;
    var _invTotal = _invExist + _invBack2 + _invExNeg2;
    if (_esVar) {
      var _pvPeso = parseFloat(o.peso) || 0;
      var _pvLong = (function(s){ var m = s.match(/[\\d.]+/); return m ? parseFloat(m[0]) : 0; })(String(o.long||''));
      var _pvFact = (_pvPeso > 0 && _pvLong > 0) ? _pvPeso * _pvLong : 0;
      if (_pvFact > 0) _restanUnid = _restanKg / _pvFact;
    }
    var _total = _invTotal + _restanUnid;
    if (_total > _invMax * 1.10) _claseInv = ' planif-sobremax';
    else if (_total < _invMax * 0.95) _claseInv = ' planif-bajomax';
  }"""
    content = content.replace(old, new, 1)
    open(path, 'w', encoding='utf-8').write(content)
    print('C1 GUARDADO')
else:
    print('C1 NO ENCONTRADO')
