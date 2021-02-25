import  pymongo
import functions

class MongoDBOP:
    def __init__(self):
        myclient = pymongo.MongoClient(functions.get_mango_client_name())
        mydb = myclient[functions.get_mango_db_name()]
        self.mycol = mydb[functions.get_mango_col_name()]

    def insert_one(self,this_dict):
        self.mycol.insert(this_dict)

    def insert_many(self,data_list):
        self.mycol.insert_many(data_list)

    def del__col(self,):
        self.mycol.delete_many({})

    def get_all(self,):
        return self.mycol.find()

    def find_sheet_fields(self,sheet_name):
        return self.mycol.find({}, {"name": 1, sheet_name: {"value": 1, "error": 1}})

    def find_cells_value(self,target_name):
        filter_data=functions.get_target_cell_pos(target_name)
        return self.mycol.find({"sheet_name":filter_data["sheet_name"],"uid":int(filter_data["row_name"])-1}, {"uid": 1
            , "sheet_name": 1,filter_data["col_name"]:1})

        pass


if __name__=="__main__":
    m=MongoDBOP()
    l=list(m.find_cells_value("Sheet1(j2)"))
    pass