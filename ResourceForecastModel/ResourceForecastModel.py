import numpy as np
import pandas as pd
import pickle
import joblib
import model
import xlwings as xw

'''

def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"


@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("ResourceForecastModel.xlsm").set_mock_caller()
    main()
'''

@xw.func
def predict_sal(subband,skill,experience):
    exp = int(experience)
    
    array = np.array([[subband,skill,exp]])
    
    index_values=[0]
    column_values=['Subband','Category','Experience']
    
    X_newTest = pd.DataFrame(data=array,
                             index=index_values,
                             columns = column_values)
    
    X_newTest = model.Band_encoder.transform(X_newTest)
    X_newTest = model.Cat_encoder.transform(X_newTest)
    
    #print(X_newTest)
    X_newTest = model.scaler.transform(X_newTest)
    
    rf_model = pickle.load(open('RF_model.pkl', 'rb'))
    
    predict = rf_model.predict(X_newTest)
    
    output = int(predict[0])
    return output