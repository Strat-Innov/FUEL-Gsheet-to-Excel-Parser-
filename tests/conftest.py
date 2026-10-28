import shutil
from pathlib import Path

import pytest

from fuel_parser.pipeline import RunOptions, run

from .fixtures.builders import TEAMS_TEST_YAML, write_source, write_template

REPO = Path(__file__).resolve().parent.parent


@pytest.fixture
def env(tmp_path):
    """A private working area: copied config with test teams, template, output, ledger."""
    cfg = tmp_path / "config"
    shutil.copytree(REPO / "config", cfg)
    (cfg / "teams.yaml").write_text(TEAMS_TEST_YAML, encoding="utf-8")
    for p in ("sellers.csv", "sellers.yaml"):
        (cfg / p).unlink(missing_ok=True)
    template = write_template(tmp_path / "template.xlsx")

    class Env:
        root = tmp_path
        config = cfg
        out = tmp_path / "out"
        ledger = tmp_path / "out" / "ledger.sqlite"
        profiles = tmp_path / "profiles"

        def source(self, rows, name="source.xlsx", **kw):
            return write_source(tmp_path / name, rows, **kw)

        def run(self, source, **kw):
            kw.setdefault("last_id", 100)
            return run(RunOptions(source=source, template=kw.pop("template", template), out_dir=self.out,
                                  config_dir=cfg, ledger_path=self.ledger, profiles_dir=self.profiles, **kw))

    Env.template = template
    return Env()
