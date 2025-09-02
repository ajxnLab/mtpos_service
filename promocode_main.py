from config.env_config import load_environment
from promo_code.promocode_service import PromoCode

if __name__ == "__main__":
    load_environment()
    service = PromoCode()
    service.run()