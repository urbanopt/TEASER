import ipdb
import os
from teaser.project import Project


def main():
    root_dir = os.path.dirname(os.path.realpath(__file__))
    file_path = os.path.join(root_dir, 'fzj_clean.gml')

    prj = Project(load_data=True)
    prj.name = 'Project_cityGML_random'
    prj.load_citygml(path=file_path)
    prj.calc_all_buildings()
    ipdb.set_trace()  # Break Point ###########

    # prj.export_aixlib(path=root_dir)
    # prj.save_project(file_name='cityGML_random', path=root_dir)
    import pickle

    pickle_file = 'teaser_pickle_only_office.p'
    pickle.dump(prj, open(pickle_file, "wb"))
    prj = pickle.load(open(pickle_file, "rb"))
    ipdb.set_trace()  # Break Point ###########

    prj.export_aixlib(path=root_dir)
    None


if __name__ == "__main__":
    main()
