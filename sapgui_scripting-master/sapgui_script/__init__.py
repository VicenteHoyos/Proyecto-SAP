import imp
import pip


def main():
    modules = ["pypiwin32", "csv"]
    for name in modules:
        try:
            imp.find_module(name)
        except:
            pip(["install",name,"--proxy","http://10.20.30.40:8080"])

if __name__ == "__main__":
    main()
    
    
