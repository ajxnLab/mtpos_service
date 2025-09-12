import sys
from config.env_config import load_environment
from matcode_mtpos.mtpos_service import Mtpos_Service

env = sys.argv[1] if len(sys.argv) > 1 else None


if __name__ == "__main__":

    if env == "dev":
        variable = "dev"
    elif env == "publish":
        variable = None
        publish = "publish"
    else:
        variable = None
        publish = None


    load_environment(variable)
    service = Mtpos_Service(publish)
    service.run()