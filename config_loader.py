import yaml
import os

CONFIG_DIR = 'config'

def load_yaml(filename):
    path = os.path.join(CONFIG_DIR, filename)
    with open(path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)

class Config:
    def __init__(self):
        self.score_weights = load_yaml('score_weights.yaml')
        self.lang_zh = load_yaml('lang_zh.yaml')
        self.lang_en = load_yaml('lang_en.yaml')
        self.general = load_yaml('config.yaml')

config = Config() 