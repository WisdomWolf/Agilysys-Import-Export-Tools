from distutils.core import setup
import py2exe

Mydata_files = [('resources', ['resources/Format_Gears.ico'])]

setup(
		windows=[
            {
                'script': 'Agilysys Import EXport Tool.pyw',
                "icon_resources": [(1, "Format_Gears.ico")]
            }
        ],
		data_files = Mydata_files,
		options={
					"py2exe": {
						"bundle_files": 2
					}
				}
)