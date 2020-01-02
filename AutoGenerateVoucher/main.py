import generater.funcs as funcs

if __name__ == "__main__":

    print("\n"+19*"="+"\n"+"Auto-Generater V1.3"+"\n"+19*"="+"\n")
    funcs.decide_generate_method(funcs.get_voucher_type())
    print("\nThe voucher template has been completed!")
    funcs.exit_with_anykey()
