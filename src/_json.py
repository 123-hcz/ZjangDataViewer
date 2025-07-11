import json
import os

def load_json(file_path):
    """
    Loads a JSON file and returns its contents as a Python object.

    Args:
        file_path (str): The path to the JSON file.

    Returns:
        object: The contents of the JSON file as a Python object.

    Raises:
        FileNotFoundError: If the specified file does not exist.
        _json.JSONDecodeError: If the JSON file is not valid.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' does not exist.")
    with open(file_path, 'r') as file:
        return json.load(file)
    return json.loads(file_path)
def save_json(file_path, data):
    """
    Saves a Python object as a JSON file.

    Args:
        file_path (str): The path to the JSON file.
        data (object): The Python object to save as JSON.

    Raises:
        FileNotFoundError: If the specified directory does not exist.
    """
    directory = os.path.dirname(file_path)
    if not os.path.exists(directory):
        os.makedirs(directory)
    with open(file_path, 'w') as file:
        json.dump(data, file, indent=4)
        print(f"Saved JSON to '{file_path}'.")

if __name__ == "__main__":
    file_path = '_json/options.json'
    data = load_json(file_path)
    print(data)

