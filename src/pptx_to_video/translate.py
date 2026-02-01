from pprint import pprint

import argostranslate
from argostranslate import package, translate

argostranslate.package.update_package_index()
available_packages = package.get_available_packages()
pprint([x.to_code for x in available_packages if x.from_code == "en"])
