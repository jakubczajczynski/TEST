from coloraide import Color
from coloraide.spaces import Cylindrical, Space
import sys

from coloraide import Color as Base
from coloraide.spaces.okhsl import Okhsl

class Color(Base): ...

# Register plugins only if not already registered
if "okhsl" not in Color.CS_MAP:
    Color.register(Okhsl())
if "okhsv" not in Color.CS_MAP:
    Color.register(Okhsv())

color = Color("okhsl", [Hue, Saturation, Lightness], 1)

X = color.convert("srgb").to_string(hex=True)  # Output as HEX