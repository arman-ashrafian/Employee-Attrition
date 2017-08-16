import numpy as np
import pandas as pd
from sklearn import preprocessing, model_selection, neighbors, svm, neural_network, tree, naive_bayes, linear_model
from sklearn.feature_extraction import DictVectorizer

def encode_onehot(df, cols):
    """
    @param df pandas DataFrame
    @param cols a list of columns to encode
    @return a DataFrame with one-hot encoding
    """
    vec = DictVectorizer()

    vec_data = pd.DataFrame(vec.fit_transform(df[cols].to_dict(outtype='records')).toarray())
    vec_data.columns = vec.get_feature_names()
    vec_data.index = df.index

    df = df.drop(cols, axis=1)
    df = df.join(vec_data)
    return df

def main():
    df = pd.read_excel('Normalized.xlsx', sheetname='Training')

    # drop employee_number
    df.drop(['EmployeeNumber'], 1, inplace=True)

    # drop over 18
    df.drop(['Over18'], 1, inplace=True)

    #drop standard hours
    df.drop(['StandardHours'], 1, inplace=True)

    # encode categorical data
    df = encode_onehot(df, cols=['BusinessTravel', 'Department', 'EducationField', 'Gender', 'JobRole', 'MaritalStatus', 'OverTime'])
    print(df.columns)
    return

    # features
    x = np.array(df.drop(['Attrition'], 1))
    # labels
    y = np.array(df['Attrition'])

    # split train/test data
    x_train, x_test, y_train, y_test = model_selection.train_test_split(x, y, test_size=0.30)

    # scale data
    scaler = preprocessing.StandardScaler()
    scaler.fit(x_train)
    x_train = scaler.transform(x_train)
    x_test = scaler.transform(x_test)

    # classifier
    clf = svm.LinearSVC()

    # clf = neural_network.MLPClassifier(hidden_layer_sizes=(10,10,100), max_iter=1000, verbose=False, tol=1e-10)


    clf.fit(x_train, y_train)

    accuracy = clf.score(x_test, y_test)

    print('\nAccuracy: %f' % accuracy)
main()
