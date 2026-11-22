// Estilos de tabla integrados de Excel, sacados del propio Excel: por cada parte de la tabla (w tabla entera,
// h encabezado, t totales, f primera columna, l última, r1/r2 y c1/c2 franjas de filas y columnas, fh/lh/ft/lt celdas
// de las esquinas), su relleno, color de fuente, negrita y bordes (l, t, b, r, v dentro vertical, h dentro horizontal).
// Un color es un índice del tema, [índice, tono] o un color fijo.
(function (g) {
  const O = (g.Oficina = g.Oficina || {});
  O.TABLE_STYLES = {
    Light1: {"w":{"color":1,"bd":{"t":["thin",1],"b":["thin",1]}},"h":{"color":1,"b":1,"bd":{"b":["thin",1]}},"t":{"color":1,"b":1,"bd":{"t":["thin",1]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[0,-0.15]},"c1":{"fill":[0,-0.15]}},
    Light2: {"w":{"color":[4,-0.25],"bd":{"t":["thin",4],"b":["thin",4]}},"h":{"color":[4,-0.25],"b":1,"bd":{"b":["thin",4]}},"t":{"color":[4,-0.25],"b":1,"bd":{"t":["thin",4]}},"f":{"color":[4,-0.25],"b":1},"l":{"color":[4,-0.25],"b":1},"r1":{"fill":[4,0.8]},"c1":{"fill":[4,0.8]}},
    Light3: {"w":{"color":[5,-0.25],"bd":{"t":["thin",5],"b":["thin",5]}},"h":{"color":[5,-0.25],"b":1,"bd":{"b":["thin",5]}},"t":{"color":[5,-0.25],"b":1,"bd":{"t":["thin",5]}},"f":{"color":[5,-0.25],"b":1},"l":{"color":[5,-0.25],"b":1},"r1":{"fill":[5,0.8]},"c1":{"fill":[5,0.8]}},
    Light4: {"w":{"color":[6,-0.25],"bd":{"t":["thin",6],"b":["thin",6]}},"h":{"color":[6,-0.25],"b":1,"bd":{"b":["thin",6]}},"t":{"color":[6,-0.25],"b":1,"bd":{"t":["thin",6]}},"f":{"color":[6,-0.25],"b":1},"l":{"color":[6,-0.25],"b":1},"r1":{"fill":[6,0.8]},"c1":{"fill":[6,0.8]}},
    Light5: {"w":{"color":[7,-0.25],"bd":{"t":["thin",7],"b":["thin",7]}},"h":{"color":[7,-0.25],"b":1,"bd":{"b":["thin",7]}},"t":{"color":[7,-0.25],"b":1,"bd":{"t":["thin",7]}},"f":{"color":[7,-0.25],"b":1},"l":{"color":[7,-0.25],"b":1},"r1":{"fill":[7,0.8]},"c1":{"fill":[7,0.8]}},
    Light6: {"w":{"color":[8,-0.25],"bd":{"t":["thin",8],"b":["thin",8]}},"h":{"color":[8,-0.25],"b":1,"bd":{"b":["thin",8]}},"t":{"color":[8,-0.25],"b":1,"bd":{"t":["thin",8]}},"f":{"color":[8,-0.25],"b":1},"l":{"color":[8,-0.25],"b":1},"r1":{"fill":[8,0.8]},"c1":{"fill":[8,0.8]}},
    Light7: {"w":{"color":[9,-0.25],"bd":{"t":["thin",9],"b":["thin",9]}},"h":{"color":[9,-0.25],"b":1,"bd":{"b":["thin",9]}},"t":{"color":[9,-0.25],"b":1,"bd":{"t":["thin",9]}},"f":{"color":[9,-0.25],"b":1},"l":{"color":[9,-0.25],"b":1},"r1":{"fill":[9,0.8]},"c1":{"fill":[9,0.8]}},
    Light8: {"w":{"color":1,"bd":{"l":["thin",1],"t":["thin",1],"b":["thin",1],"r":["thin",1]}},"h":{"fill":1,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",1]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"bd":{"t":["thin",1]}},"r2":{"bd":{"t":["thin",1]}},"c1":{"bd":{"l":["thin",1]}},"c2":{"bd":{"l":["thin",1]}}},
    Light9: {"w":{"color":1,"bd":{"l":["thin",4],"t":["thin",4],"b":["thin",4],"r":["thin",4]}},"h":{"fill":4,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",4]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"bd":{"t":["thin",4]}},"r2":{"bd":{"t":["thin",4]}},"c1":{"bd":{"l":["thin",4]}},"c2":{"bd":{"l":["thin",4]}}},
    Light10: {"w":{"color":1,"bd":{"l":["thin",5],"t":["thin",5],"b":["thin",5],"r":["thin",5]}},"h":{"fill":5,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",5]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"bd":{"t":["thin",5]}},"r2":{"bd":{"t":["thin",5]}},"c1":{"bd":{"l":["thin",5]}},"c2":{"bd":{"l":["thin",5]}}},
    Light11: {"w":{"color":1,"bd":{"l":["thin",6],"t":["thin",6],"b":["thin",6],"r":["thin",6]}},"h":{"fill":6,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",6]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"bd":{"t":["thin",6]}},"r2":{"bd":{"t":["thin",6]}},"c1":{"bd":{"l":["thin",6]}},"c2":{"bd":{"l":["thin",6]}}},
    Light12: {"w":{"color":1,"bd":{"l":["thin",7],"t":["thin",7],"b":["thin",7],"r":["thin",7]}},"h":{"fill":7,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",7]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"bd":{"t":["thin",7]}},"r2":{"bd":{"t":["thin",7]}},"c1":{"bd":{"l":["thin",7]}},"c2":{"bd":{"l":["thin",7]}}},
    Light13: {"w":{"color":1,"bd":{"l":["thin",8],"t":["thin",8],"b":["thin",8],"r":["thin",8]}},"h":{"fill":8,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",8]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"bd":{"t":["thin",8]}},"r2":{"bd":{"t":["thin",8]}},"c1":{"bd":{"l":["thin",8]}},"c2":{"bd":{"l":["thin",8]}}},
    Light14: {"w":{"color":1,"bd":{"l":["thin",9],"t":["thin",9],"b":["thin",9],"r":["thin",9]}},"h":{"fill":9,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",9]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"bd":{"t":["thin",9]}},"r2":{"bd":{"t":["thin",9]}},"c1":{"bd":{"l":["thin",9]}},"c2":{"bd":{"l":["thin",9]}}},
    Light15: {"w":{"color":1,"bd":{"l":["thin",1],"t":["thin",1],"b":["thin",1],"r":["thin",1],"v":["thin",1],"h":["thin",1]}},"h":{"color":1,"b":1,"bd":{"b":["medium",1]}},"t":{"color":1,"b":1,"bd":{"t":["double",1]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[0,-0.15]},"c1":{"fill":[0,-0.15]}},
    Light16: {"w":{"color":1,"bd":{"l":["thin",4],"t":["thin",4],"b":["thin",4],"r":["thin",4],"v":["thin",4],"h":["thin",4]}},"h":{"color":1,"b":1,"bd":{"b":["medium",4]}},"t":{"color":1,"b":1,"bd":{"t":["double",4]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[4,0.8]},"c1":{"fill":[4,0.8]}},
    Light17: {"w":{"color":1,"bd":{"l":["thin",5],"t":["thin",5],"b":["thin",5],"r":["thin",5],"v":["thin",5],"h":["thin",5]}},"h":{"color":1,"b":1,"bd":{"b":["medium",5]}},"t":{"color":1,"b":1,"bd":{"t":["double",5]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[5,0.8]},"c1":{"fill":[5,0.8]}},
    Light18: {"w":{"color":1,"bd":{"l":["thin",6],"t":["thin",6],"b":["thin",6],"r":["thin",6],"v":["thin",6],"h":["thin",6]}},"h":{"color":1,"b":1,"bd":{"b":["medium",6]}},"t":{"color":1,"b":1,"bd":{"t":["double",6]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[6,0.8]},"c1":{"fill":[6,0.8]}},
    Light19: {"w":{"color":1,"bd":{"l":["thin",7],"t":["thin",7],"b":["thin",7],"r":["thin",7],"v":["thin",7],"h":["thin",7]}},"h":{"color":1,"b":1,"bd":{"b":["medium",7]}},"t":{"color":1,"b":1,"bd":{"t":["double",7]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[7,0.8]},"c1":{"fill":[7,0.8]}},
    Light20: {"w":{"color":1,"bd":{"l":["thin",8],"t":["thin",8],"b":["thin",8],"r":["thin",8],"v":["thin",8],"h":["thin",8]}},"h":{"color":1,"b":1,"bd":{"b":["medium",8]}},"t":{"color":1,"b":1,"bd":{"t":["double",8]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[8,0.8]},"c1":{"fill":[8,0.8]}},
    Light21: {"w":{"color":1,"bd":{"l":["thin",9],"t":["thin",9],"b":["thin",9],"r":["thin",9],"v":["thin",9],"h":["thin",9]}},"h":{"color":1,"b":1,"bd":{"b":["medium",9]}},"t":{"color":1,"b":1,"bd":{"t":["double",9]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[9,0.8]},"c1":{"fill":[9,0.8]}},
    Medium1: {"w":{"color":1,"bd":{"l":["thin",1],"t":["thin",1],"b":["thin",1],"r":["thin",1],"h":["thin",1]}},"h":{"fill":1,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",1]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[0,-0.15]},"c1":{"fill":[0,-0.15]}},
    Medium2: {"w":{"color":1,"bd":{"l":["thin",[4,0.4]],"t":["thin",[4,0.4]],"b":["thin",[4,0.4]],"r":["thin",[4,0.4]],"h":["thin",[4,0.4]]}},"h":{"fill":4,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",4]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[4,0.8]},"c1":{"fill":[4,0.8]}},
    Medium3: {"w":{"color":1,"bd":{"l":["thin",[5,0.4]],"t":["thin",[5,0.4]],"b":["thin",[5,0.4]],"r":["thin",[5,0.4]],"h":["thin",[5,0.4]]}},"h":{"fill":5,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",5]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[5,0.8]},"c1":{"fill":[5,0.8]}},
    Medium4: {"w":{"color":1,"bd":{"l":["thin",[6,0.4]],"t":["thin",[6,0.4]],"b":["thin",[6,0.4]],"r":["thin",[6,0.4]],"h":["thin",[6,0.4]]}},"h":{"fill":6,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",6]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[6,0.8]},"c1":{"fill":[6,0.8]}},
    Medium5: {"w":{"color":1,"bd":{"l":["thin",[7,0.4]],"t":["thin",[7,0.4]],"b":["thin",[7,0.4]],"r":["thin",[7,0.4]],"h":["thin",[7,0.4]]}},"h":{"fill":7,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",7]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[7,0.8]},"c1":{"fill":[7,0.8]}},
    Medium6: {"w":{"color":1,"bd":{"l":["thin",[8,0.4]],"t":["thin",[8,0.4]],"b":["thin",[8,0.4]],"r":["thin",[8,0.4]],"h":["thin",[8,0.4]]}},"h":{"fill":8,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",8]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[8,0.8]},"c1":{"fill":[8,0.8]}},
    Medium7: {"w":{"color":1,"bd":{"l":["thin",[9,0.4]],"t":["thin",[9,0.4]],"b":["thin",[9,0.4]],"r":["thin",[9,0.4]],"h":["thin",[9,0.4]]}},"h":{"fill":9,"color":0,"b":1},"t":{"color":1,"b":1,"bd":{"t":["double",9]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[9,0.8]},"c1":{"fill":[9,0.8]}},
    Medium8: {"w":{"fill":[0,-0.15],"color":1,"bd":{"v":["thin",0],"h":["thin",0]}},"h":{"fill":1,"color":0,"b":1,"bd":{"b":["thick",0]}},"t":{"fill":1,"color":0,"b":1,"bd":{"t":["thick",0]}},"f":{"fill":1,"color":0,"b":1},"l":{"fill":1,"color":0,"b":1},"r1":{"fill":[0,-0.35]},"c1":{"fill":[0,-0.35]}},
    Medium9: {"w":{"fill":[4,0.8],"color":1,"bd":{"v":["thin",0],"h":["thin",0]}},"h":{"fill":4,"color":0,"b":1,"bd":{"b":["thick",0]}},"t":{"fill":4,"color":0,"b":1,"bd":{"t":["thick",0]}},"f":{"fill":4,"color":0,"b":1},"l":{"fill":4,"color":0,"b":1},"r1":{"fill":[4,0.6]},"c1":{"fill":[4,0.6]}},
    Medium10: {"w":{"fill":[5,0.8],"color":1,"bd":{"v":["thin",0],"h":["thin",0]}},"h":{"fill":5,"color":0,"b":1,"bd":{"b":["thick",0]}},"t":{"fill":5,"color":0,"b":1,"bd":{"t":["thick",0]}},"f":{"fill":5,"color":0,"b":1},"l":{"fill":5,"color":0,"b":1},"r1":{"fill":[5,0.6]},"c1":{"fill":[5,0.6]}},
    Medium11: {"w":{"fill":[6,0.8],"color":1,"bd":{"v":["thin",0],"h":["thin",0]}},"h":{"fill":6,"color":0,"b":1,"bd":{"b":["thick",0]}},"t":{"fill":6,"color":0,"b":1,"bd":{"t":["thick",0]}},"f":{"fill":6,"color":0,"b":1},"l":{"fill":6,"color":0,"b":1},"r1":{"fill":[6,0.6]},"c1":{"fill":[6,0.6]}},
    Medium12: {"w":{"fill":[7,0.8],"color":1,"bd":{"v":["thin",0],"h":["thin",0]}},"h":{"fill":7,"color":0,"b":1,"bd":{"b":["thick",0]}},"t":{"fill":7,"color":0,"b":1,"bd":{"t":["thick",0]}},"f":{"fill":7,"color":0,"b":1},"l":{"fill":7,"color":0,"b":1},"r1":{"fill":[7,0.6]},"c1":{"fill":[7,0.6]}},
    Medium13: {"w":{"fill":[8,0.8],"color":1,"bd":{"v":["thin",0],"h":["thin",0]}},"h":{"fill":8,"color":0,"b":1,"bd":{"b":["thick",0]}},"t":{"fill":8,"color":0,"b":1,"bd":{"t":["thick",0]}},"f":{"fill":8,"color":0,"b":1},"l":{"fill":8,"color":0,"b":1},"r1":{"fill":[8,0.6]},"c1":{"fill":[8,0.6]}},
    Medium14: {"w":{"fill":[9,0.8],"color":1,"bd":{"v":["thin",0],"h":["thin",0]}},"h":{"fill":9,"color":0,"b":1,"bd":{"b":["thick",0]}},"t":{"fill":9,"color":0,"b":1,"bd":{"t":["thick",0]}},"f":{"fill":9,"color":0,"b":1},"l":{"fill":9,"color":0,"b":1},"r1":{"fill":[9,0.6]},"c1":{"fill":[9,0.6]}},
    Medium15: {"w":{"color":1,"bd":{"l":["thin",1],"t":["medium",1],"b":["medium",1],"r":["thin",1],"v":["thin",1],"h":["thin",1]}},"h":{"fill":1,"color":0,"b":1,"bd":{"b":["medium",1]}},"t":{"bd":{"t":["double",1]}},"f":{"fill":1,"color":0,"b":1},"l":{"fill":1,"color":0,"b":1},"r1":{"fill":[0,-0.15]},"c1":{"fill":[0,-0.15]}},
    Medium16: {"w":{"color":1,"bd":{"t":["medium",1],"b":["medium",1]}},"h":{"fill":4,"color":0,"b":1,"bd":{"b":["medium",1]}},"t":{"bd":{"t":["double",1]}},"f":{"fill":4,"color":0,"b":1},"l":{"fill":4,"color":0,"b":1},"r1":{"fill":[0,-0.15]},"c1":{"fill":[0,-0.15]}},
    Medium17: {"w":{"color":1,"bd":{"t":["medium",1],"b":["medium",1]}},"h":{"fill":5,"color":0,"b":1,"bd":{"b":["medium",1]}},"t":{"bd":{"t":["double",1]}},"f":{"fill":5,"color":0,"b":1},"l":{"fill":5,"color":0,"b":1},"r1":{"fill":[0,-0.15]},"c1":{"fill":[0,-0.15]}},
    Medium18: {"w":{"color":1,"bd":{"t":["medium",1],"b":["medium",1]}},"h":{"fill":6,"color":0,"b":1,"bd":{"b":["medium",1]}},"t":{"bd":{"t":["double",1]}},"f":{"fill":6,"color":0,"b":1},"l":{"fill":6,"color":0,"b":1},"r1":{"fill":[0,-0.15]},"c1":{"fill":[0,-0.15]}},
    Medium19: {"w":{"color":1,"bd":{"t":["medium",1],"b":["medium",1]}},"h":{"fill":7,"color":0,"b":1,"bd":{"b":["medium",1]}},"t":{"bd":{"t":["double",1]}},"f":{"fill":7,"color":0,"b":1},"l":{"fill":7,"color":0,"b":1},"r1":{"fill":[0,-0.15]},"c1":{"fill":[0,-0.15]}},
    Medium20: {"w":{"color":1,"bd":{"t":["medium",1],"b":["medium",1]}},"h":{"fill":8,"color":0,"b":1,"bd":{"b":["medium",1]}},"t":{"bd":{"t":["double",1]}},"f":{"fill":8,"color":0,"b":1},"l":{"fill":8,"color":0,"b":1},"r1":{"fill":[0,-0.15]},"c1":{"fill":[0,-0.15]}},
    Medium21: {"w":{"color":1,"bd":{"t":["medium",1],"b":["medium",1]}},"h":{"fill":9,"color":0,"b":1,"bd":{"b":["medium",1]}},"t":{"bd":{"t":["double",1]}},"f":{"fill":9,"color":0,"b":1},"l":{"fill":9,"color":0,"b":1},"r1":{"fill":[0,-0.15]},"c1":{"fill":[0,-0.15]}},
    Medium22: {"w":{"fill":[0,-0.15],"color":1,"bd":{"l":["thin",1],"t":["thin",1],"b":["thin",1],"r":["thin",1],"v":["thin",1],"h":["thin",1]}},"h":{"color":1,"b":1},"t":{"color":1,"b":1,"bd":{"t":["medium",1]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[0,-0.35]},"c1":{"fill":[0,-0.35]}},
    Medium23: {"w":{"fill":[4,0.8],"color":1,"bd":{"l":["thin",[4,0.4]],"t":["thin",[4,0.4]],"b":["thin",[4,0.4]],"r":["thin",[4,0.4]],"v":["thin",[4,0.4]],"h":["thin",[4,0.4]]}},"h":{"color":1,"b":1},"t":{"color":1,"b":1,"bd":{"t":["medium",4]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[4,0.6]},"c1":{"fill":[4,0.6]}},
    Medium24: {"w":{"fill":[5,0.8],"color":1,"bd":{"l":["thin",[5,0.4]],"t":["thin",[5,0.4]],"b":["thin",[5,0.4]],"r":["thin",[5,0.4]],"v":["thin",[5,0.4]],"h":["thin",[5,0.4]]}},"h":{"color":1,"b":1},"t":{"color":1,"b":1,"bd":{"t":["medium",5]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[5,0.6]},"c1":{"fill":[5,0.6]}},
    Medium25: {"w":{"fill":[6,0.8],"color":1,"bd":{"l":["thin",[6,0.4]],"t":["thin",[6,0.4]],"b":["thin",[6,0.4]],"r":["thin",[6,0.4]],"v":["thin",[6,0.4]],"h":["thin",[6,0.4]]}},"h":{"color":1,"b":1},"t":{"color":1,"b":1,"bd":{"t":["medium",6]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[6,0.6]},"c1":{"fill":[6,0.6]}},
    Medium26: {"w":{"fill":[7,0.8],"color":1,"bd":{"l":["thin",[7,0.4]],"t":["thin",[7,0.4]],"b":["thin",[7,0.4]],"r":["thin",[7,0.4]],"v":["thin",[7,0.4]],"h":["thin",[7,0.4]]}},"h":{"color":1,"b":1},"t":{"color":1,"b":1,"bd":{"t":["medium",7]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[7,0.6]},"c1":{"fill":[7,0.6]}},
    Medium27: {"w":{"fill":[8,0.8],"color":1,"bd":{"l":["thin",[8,0.4]],"t":["thin",[8,0.4]],"b":["thin",[8,0.4]],"r":["thin",[8,0.4]],"v":["thin",[8,0.4]],"h":["thin",[8,0.4]]}},"h":{"color":1,"b":1},"t":{"color":1,"b":1,"bd":{"t":["medium",8]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[8,0.6]},"c1":{"fill":[8,0.6]}},
    Medium28: {"w":{"fill":[9,0.8],"color":1,"bd":{"l":["thin",[9,0.4]],"t":["thin",[9,0.4]],"b":["thin",[9,0.4]],"r":["thin",[9,0.4]],"v":["thin",[9,0.4]],"h":["thin",[9,0.4]]}},"h":{"color":1,"b":1},"t":{"color":1,"b":1,"bd":{"t":["medium",9]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[9,0.6]},"c1":{"fill":[9,0.6]}},
    Dark1: {"w":{"fill":[1,0.45],"color":0},"h":{"fill":1,"color":0,"b":1,"bd":{"b":["medium",0]}},"t":{"fill":[1,0.15],"color":0,"b":1,"bd":{"t":["medium",0]}},"f":{"fill":[1,0.25],"color":0,"b":1,"bd":{"r":["medium",0]}},"l":{"fill":[1,0.25],"color":0,"b":1,"bd":{"l":["medium",0]}},"r1":{"fill":[1,0.25]},"c1":{"fill":[1,0.25]}},
    Dark2: {"w":{"fill":4,"color":0},"h":{"fill":1,"color":0,"b":1,"bd":{"b":["medium",0]}},"t":{"fill":[4,-0.5],"color":0,"b":1,"bd":{"t":["medium",0]}},"f":{"fill":[4,-0.25],"color":0,"b":1,"bd":{"r":["medium",0]}},"l":{"fill":[4,-0.25],"color":0,"b":1,"bd":{"l":["medium",0]}},"r1":{"fill":[4,-0.25]},"c1":{"fill":[4,-0.25]}},
    Dark3: {"w":{"fill":5,"color":0},"h":{"fill":1,"color":0,"b":1,"bd":{"b":["medium",0]}},"t":{"fill":[5,-0.5],"color":0,"b":1,"bd":{"t":["medium",0]}},"f":{"fill":[5,-0.25],"color":0,"b":1,"bd":{"r":["medium",0]}},"l":{"fill":[5,-0.25],"color":0,"b":1,"bd":{"l":["medium",0]}},"r1":{"fill":[5,-0.25]},"c1":{"fill":[5,-0.25]}},
    Dark4: {"w":{"fill":6,"color":0},"h":{"fill":1,"color":0,"b":1,"bd":{"b":["medium",0]}},"t":{"fill":[6,-0.5],"color":0,"b":1,"bd":{"t":["medium",0]}},"f":{"fill":[6,-0.25],"color":0,"b":1,"bd":{"r":["medium",0]}},"l":{"fill":[6,-0.25],"color":0,"b":1,"bd":{"l":["medium",0]}},"r1":{"fill":[6,-0.25]},"c1":{"fill":[6,-0.25]}},
    Dark5: {"w":{"fill":7,"color":0},"h":{"fill":1,"color":0,"b":1,"bd":{"b":["medium",0]}},"t":{"fill":[7,-0.5],"color":0,"b":1,"bd":{"t":["medium",0]}},"f":{"fill":[7,-0.25],"color":0,"b":1,"bd":{"r":["medium",0]}},"l":{"fill":[7,-0.25],"color":0,"b":1,"bd":{"l":["medium",0]}},"r1":{"fill":[7,-0.25]},"c1":{"fill":[7,-0.25]}},
    Dark6: {"w":{"fill":8,"color":0},"h":{"fill":1,"color":0,"b":1,"bd":{"b":["medium",0]}},"t":{"fill":[8,-0.5],"color":0,"b":1,"bd":{"t":["medium",0]}},"f":{"fill":[8,-0.25],"color":0,"b":1,"bd":{"r":["medium",0]}},"l":{"fill":[8,-0.25],"color":0,"b":1,"bd":{"l":["medium",0]}},"r1":{"fill":[8,-0.25]},"c1":{"fill":[8,-0.25]}},
    Dark7: {"w":{"fill":9,"color":0},"h":{"fill":1,"color":0,"b":1,"bd":{"b":["medium",0]}},"t":{"fill":[9,-0.5],"color":0,"b":1,"bd":{"t":["medium",0]}},"f":{"fill":[9,-0.25],"color":0,"b":1,"bd":{"r":["medium",0]}},"l":{"fill":[9,-0.25],"color":0,"b":1,"bd":{"l":["medium",0]}},"r1":{"fill":[9,-0.25]},"c1":{"fill":[9,-0.25]}},
    Dark8: {"w":{"fill":[0,-0.15]},"h":{"fill":1,"color":0},"t":{"color":1,"b":1,"bd":{"t":["double",1]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[0,-0.35]},"c1":{"fill":[0,-0.35]}},
    Dark9: {"w":{"fill":[4,0.8]},"h":{"fill":5,"color":0},"t":{"color":1,"b":1,"bd":{"t":["double",1]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[4,0.6]},"c1":{"fill":[4,0.6]}},
    Dark10: {"w":{"fill":[6,0.8]},"h":{"fill":7,"color":0},"t":{"color":1,"b":1,"bd":{"t":["double",1]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[6,0.6]},"c1":{"fill":[6,0.6]}},
    Dark11: {"w":{"fill":[8,0.8]},"h":{"fill":9,"color":0},"t":{"color":1,"b":1,"bd":{"t":["double",1]}},"f":{"color":1,"b":1},"l":{"color":1,"b":1},"r1":{"fill":[8,0.6]},"c1":{"fill":[8,0.6]}},
  };

  // Tema de Office con el que están pensados los estilos (lt1, dk1, lt2, dk2, acentos, vínculos).
  const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
  const hex = (c, theme) => {
    if (typeof c === 'string') return c;
    const [i, tint] = Array.isArray(c) ? c : [c, 0];
    const base = (theme || THEME)[i] || THEME[i];
    return tint && O.applyTint ? O.applyTint(base, tint) : base;
  };
  // Formato que pone el estilo de una tabla en una celda: relleno, fuente, negrita y bordes.
  // Como en Excel, cada parte pisa a la anterior: tabla, franjas de columnas, de filas, primera y última
  // columna, encabezado, totales y las celdas de las esquinas.
  // custom son los estilos personalizados del libro (state.tableStyles), que se buscan antes que los integrados.
  O.tableStyleAt = (sheet, r, c, theme, custom) => {
    const t = O.tableAt && O.tableAt(sheet, r, c);
    const own = t && t.style && custom && custom.list.find((x) => x.name === t.style.name);
    const def = own ? own.def : t && t.style && O.TABLE_STYLES[String(t.style.name).replace(/^TableStyle/, '')];
    if (!def) return null;
    const b = O.tableBox(t), st = t.style;
    const hdr = t.header !== false, tot = t.totals || 0;
    const d1 = b.r1 + (hdr ? 1 : 0), d2 = b.r2 - tot;
    const inData = r >= d1 && r <= d2;
    const parts = [['w', b]];
    const band = (el1, el2, at, start, end, lo, hi, box) => {
      const n1 = (def[el1] && def[el1].size) || 1, n2 = (def[el2] && def[el2].size) || 1;
      const k = (at - start) % (n1 + n2);
      const first = k < n1;
      const from = at - (first ? k : k - n1);
      const len = first ? n1 : n2;
      parts.push([first ? el1 : el2, box(from, Math.min(end, from + len - 1))]);
    };
    if (st.cols && inData) band('c1', 'c2', c, b.c1, b.c2, d1, d2, (x, y) => ({ r1: d1, r2: d2, c1: x, c2: y }));
    if (st.rows && inData) band('r1', 'r2', r, d1, d2, b.c1, b.c2, (x, y) => ({ r1: x, r2: y, c1: b.c1, c2: b.c2 }));
    if (st.first && c === b.c1) parts.push(['f', { r1: b.r1, r2: b.r2, c1: c, c2: c }]);
    if (st.last && c === b.c2) parts.push(['l', { r1: b.r1, r2: b.r2, c1: c, c2: c }]);
    if (hdr && r === b.r1) {
      parts.push(['h', { r1: r, r2: r, c1: b.c1, c2: b.c2 }]);
      if (st.first && c === b.c1) parts.push(['fh', { r1: r, r2: r, c1: c, c2: c }]);
      if (st.last && c === b.c2) parts.push(['lh', { r1: r, r2: r, c1: c, c2: c }]);
    }
    if (tot && r > d2) {
      parts.push(['t', { r1: d2 + 1, r2: b.r2, c1: b.c1, c2: b.c2 }]);
      if (st.first && c === b.c1) parts.push(['ft', { r1: r, r2: r, c1: c, c2: c }]);
      if (st.last && c === b.c2) parts.push(['lt', { r1: r, r2: r, c1: c, c2: c }]);
    }
    const out = {}, bdx = {};
    for (const [key, box] of parts) {
      const e = def[key];
      if (!e) continue;
      if (e.fill !== undefined) out.fill = hex(e.fill, theme);
      if (e.color !== undefined) out.color = hex(e.color, theme);
      if (e.b) out.b = true;
      if (!e.bd) continue;
      // Cada lado de la celda es borde exterior de la parte o línea interior.
      const side = (k, edge, inside, isEdge) => {
        const line = e.bd[isEdge ? edge : inside];
        if (line) bdx[k] = [line[0], hex(line[1], theme)];
      };
      side('t', 't', 'h', r === box.r1);
      side('b', 'b', 'h', r === box.r2);
      side('l', 'l', 'v', c === box.c1);
      side('r', 'r', 'v', c === box.c2);
    }
    if (out.color === '000000') delete out.color;
    const sides = Object.keys(bdx);
    if (sides.length) { out.bd = sides.join(''); out.bdx = bdx; }
    return out;
  };
})(globalThis);
