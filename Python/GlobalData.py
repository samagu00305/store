# # 글로벌 변수처럼 사용할 함수
# # 이유 : 1. 글로벌 변수는 F12로 정의되어진 곳에 바로가는 기능이 안되서
# # 이유 : 2. F2 로 이름 변경이 함수는 한번에 되기 때문에


def g_StartInternationalDeliveryMessage():
    return "해외배송시작이 되었습니다. 고객님.`n`n(배송 상태가 직접전달로 된 상태에서는 해외 중간 배송지로 배송 및 항공 배송 중입니다.)`n`n통관 시작 시와 국내 배송 시작 시 문자로 드리겠습니다.`n`n궁금한 것 있으면 언제든지 문의하시면 됩니다.`n`n이용해 주셔서 감사합니다."


def g_StartCustomsClearanceMessage():
    return "통관시작 되었습니다.`n`n통관 후 국내 배송이 시작됩니다.`n`n(기본 통관 기간은 1~2일 입니다.)"


def g_StartKoreaDeliveryMessage():
    # result := "첫 번째 값: {0}, 두 번째 값: {1}".Format(firstValue, secondValue)
    return "국내배송이 시작 되었습니다.`n`n구매해주신 곳에서 송장번호로 배송 상황을 보실 수 있습니다.`n`n이용해 주셔서 감사합니다."


# 택배 가격
def g_CourierPrice():
    return 25000


# 마진율
def g_MarginRate():
    return 0.1  # 10%
