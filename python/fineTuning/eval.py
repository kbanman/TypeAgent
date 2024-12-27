# Copyright (c) Microsoft Corporation and Henry Lucco.
# Licensed under the MIT License.

from chaparral.util.datareader import DataReader
from chaparral.train.hf_model import HFModel
from chaparral.train.hf_params import HFParams
import argparse


def parse_args():
    parser = argparse.ArgumentParser(
        description="Fine-tune a model with given dataset.")
    parser.add_argument("--dataset_file", help="Path to the dataset file.")
    parser.add_argument("--model_name", help="Name of the model to fine-tune.")
    parser.add_argument("--params", help="Path to params file")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    dataset_file = args.dataset_file
    params_file = args.params

    # load params
    params = HFParams.from_file(params_file)

    # load dataset
    dataset = DataReader().load_text_file(dataset_file)

    # format data into train and eval sets
    train_set, eval_set = dataset.create_train_eval_sets()

    model = HFModel(params)

    print("Model loaded")

    model.load_local_model("./test_output")
    print(model.evaluate(eval_set))
    print(model.generate(dataset.get_filled_prompt(
        "The quick brown fox jumps over the lazy dog")))