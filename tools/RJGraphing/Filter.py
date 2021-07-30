import numpy as np
import pandas as pd
import datetime
import time
from RJGraphing import Graphing


class Branch(dict):

    def __init__(self, data, target, index, column, max_iter, address, *args):
        """
        :param data: pd.DataFrame, the input data, column 1: id, column 2: Policy Intensity
        :param target: pd.Series, the input data, index: Year, column 1: Percentage of Policy Intensity in each year
        :param loss: list, record of the ols_loss in the latest screening
        :param trail: list, which contains tuple, e.g. [(diff_max_key, diff_max, ols_loss_min_key, //
        ols_loss_min)], marking the process of fitting
        :param index: str, the name of the column that contains the index
        :param column: str, the name of the column to be summarized
        :param max_iter: int, the maximum of iterations of the fitting
        :param *args: dict, the latest status of samples, {Year1: [Sample1 id, Sample2 id], Year2: []}
        """
        self.time_start = time.time()
        self.time_end = None

        dict.__init__(self, *args)
        self.data = data
        self.target = target
        self.loss_min = 1
        self.loss_lst = []
        self.trail = []
        self.index = index
        self.column = column
        self.max_iter = max_iter
        self.actual_iter = 0
        self.address = address
        self.diff_max = None
        self.diff_max_key = None
        self.drop_sample = []

        grouped_data = data.groupby(by=index, axis=0).sum()[column]
        grouped_data = pd.DataFrame([x / grouped_data.sum() for x in grouped_data], index=grouped_data.index)
        self.diff_screening(grouped_data)

    def trail_update(self, diff_max_key, diff_max, ols_min_key, ols_min):
        new_trail = (diff_max_key, diff_max, ols_min_key, ols_min)
        self.trail.append(new_trail)

    def ols_loss(self, df_ratio):
        """
        :param df_ratio: pd.Series, the primary data(ratio of primary value)
        :param target: pd.Series, the desired ratio of each value
        :return: ols_loss: numpy.intc, the loss calculated by ordinary least squares
        """
        df_ratio = np.array(df_ratio)
        target = self.target

        if len(target) == len(df_ratio):
            pass
        else:
            raise ValueError("The target and the data do not match in length")

        diff = df_ratio - target
        diff = diff ** 2
        self.loss_lst.append(diff.sum())

    def branch_kill(self):
        self.save_data_plot()
        self.time_end = time.time()
        print('--------------------------------------------------------------------')
        print('The training has ended. Time cost: {:.4f} s'.format(self.time_end-self.time_start))
        print('--------------------------------------------------------------------')
        print('Fitting Times: ' + str(self.actual_iter))
        print('Num of Samples: ' + str(len(self.data)))
        print('Current OLS Loss: ' + str(self.loss_min))
        print('Current Diff_max: {:.2f}%'.format(100*self.loss_min))

        print('Dropped Samples: ')
        print(self.drop_sample)

        dd = dict()
        for k, v in self.items():
            dd[k] = len(self[k])
        print('Current Samples: ')
        print(dd)

        print('The Fitting Process: (diff_max_key, diff_max, ols_loss_min_key, ols_loss_min)')
        for item in self.trail:
            print(item)

    def diff_screening(self, grouped_data):
        # Diff Screening
        diff_keys = list(grouped_data.index)
        df = np.array(grouped_data)
        target = self.target

        diff = df - target
        self.diff_max = np.max(diff)
        self.diff_max_key = diff_keys[diff.argmax()]

        # Continue the new round
        self.loss_screening(self.data,
                            self.diff_max_key,
                            self.diff_max,
                            self.target,
                            self.index,
                            self.column)

    def loss_screening(self, data, diff_max_key, diff_max, target, indi, column):
        # Loss Screening
        sample = self[diff_max_key]
        for key in sample:
            df = data.copy()
            df.drop(df[df['id'] == key].index, axis=0, inplace=True)

            # Data Aggregation by summing
            df = df.groupby(by=indi, axis=0).sum()[column]
            df_ratio = pd.DataFrame([x / df.sum() for x in df], index=df.index)
            self.ols_loss(df_ratio)

        ols_min = min(self.loss_lst)
        ols_min_key = sample[self.loss_lst.index(ols_min)]

        # The first ending condition: Max Iteration
        if self.actual_iter >= self.max_iter:
            self.branch_kill()
        else:
            # The second ending condition: Convergence
            if ols_min <= self.loss_min:
                data = self.data
                data.drop(data[data['id'] == ols_min_key].index, axis=0, inplace=True)
                df = data.groupby(by=indi, axis=0).sum()[column]
                df_ratio = pd.DataFrame([x / df.sum() for x in df], index=df.index)
                self.actual_iter += 1
                self.data = data
                self.trail_update(diff_max_key, diff_max, ols_min_key, ols_min)
                self.drop_sample.append(ols_min_key)
                self.loss_min = ols_min
                self.loss_lst = []

                # Start the next round
                self.diff_screening(df_ratio)
            else:
                self.branch_kill()

    def save_data_plot(self):
        Graphing.policy_intensity(self.data, index='Year', column='pmg强度', how='Sum')
        time_now = datetime.datetime.today()
        time_now = time_now.strftime('%Y%m%d_%H%M')
        address = self.address + time_now + '.xlsx'
        self.data.to_excel(address)

