set dutil = new DictUtil
set d = argsDict
call dutil.SortDictionary(d, 1)
EchoX "Parameter Keys: %x", join(d.Keys, ",")
EchoX "Parameter Items: %x", join(d.Items, ",")
call dutil.SortDictionary(d, 2)
EchoX "Parameter Keys: %x", join(d.Keys, ",")
EchoX "Parameter Items: %x", join(d.Items, ",")