# @GHInput: center (Point3d) 
# @GHInput: base_radius (float) 
# @GHInput: random_domain (Interval) 
# @GHOutput: circle (Circle) 
# @GHOutput: final_radius (float) 

import Rhino.Geometry as rg
import random

if center is not None and base_radius is not None and random_domain is not None:
    random_offset = random.uniform(random_domain.Min, random_domain.Max)
    final_radius = base_radius + random_offset
    circle = rg.Circle(center, final_radius)
else:
    circle = None
    final_radius = None