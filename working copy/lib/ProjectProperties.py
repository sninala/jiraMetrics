import random


class Singleton(type):
    _instances = {}

    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super(Singleton, cls).__call__(*args, **kwargs)
        return cls._instances[cls]


class ProjectProperties(object):
    __metaclass__ = Singleton

    def __init__(self, config):
        self.config = config
        self.project_properties = dict()

    def initialize_project_properties(self):
        ert_projects = self.get_project_codes()
        color_section = 'PROJECT_COLOR'
        marker_section = 'PROJECT_MARKER_SYMBOL'
        colors = dict(self.config.items(color_section))
        marker_symbols = dict(self.config.items(marker_section))
        for project in ert_projects:
            (project_color, marker) = (None, None)
            if project not in colors:
                project_color = self.get_random_color_code()
                self.config.set(color_section, project, project_color)
            else:
                project_color = self.config.get(color_section, project)
            if project not in marker_symbols:
                marker = self.get_random_marker_for_project()
                self.config.set(marker_section, project, marker)
            else:
                marker = self.config.get(marker_section, project)
            self.project_properties[project] = {"MARKER_SYMBOL": marker, "COLOR": project_color }

    def get_project_properties_for(self, project):
        return self.project_properties[project]

    def get_project_codes(self):
        project_codes = [project.strip() for project in self.config.get('BUG_TRACKER', 'projects').split(",")]
        return project_codes

    @staticmethod
    def get_random_color_code():
        hex_digits = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F']
        digit_array = []
        for i in range(6):
            digit_array.append(hex_digits[random.randint(0, 15)])
        joined_digits = ''.join(digit_array)
        return joined_digits

    @staticmethod
    def get_random_marker_for_project():
        markers = ['circle', 'dash', 'diamond', 'dot', 'plus', 'square', 'star', 'triangle', 'x']
        return random.choice(markers)

