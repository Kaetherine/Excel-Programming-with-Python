import xlwings as xw


def main():
    wb = xw.Book.caller()
    base_data = xw.Book('sample_data.xlsm')
    customers = base_data.sheets[0]


@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("app.xlsm").set_mock_caller()
    main()
