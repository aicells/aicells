import importlib
import os

class AICFactory:
    classTypeCache = {}
    aicConfig = {}

    def __init__(self, aicConfig):
        self.aicConfig = aicConfig

    def CreateInstance(self, classNameWithNamespace, args={}):
        # directories = [
        #     'core',
        #     'correlation',
        #     'random',
        #     'supervised-learning',
        # ]


        #classNameWithNamespace = classNameWithNamespace

        if classNameWithNamespace in self.classTypeCache:
            classType = self.classTypeCache[classNameWithNamespace]
        else:
            classNameWithNamespaceAsList = classNameWithNamespace.split('.')
            className = classNameWithNamespaceAsList.pop()

            namespace = ''
            if len(classNameWithNamespaceAsList) != 0:
                # namespace = '.'+'.'.join(classNameWithNamespaceAsList)
                namespace = '.'.join(classNameWithNamespaceAsList)

            #module = importlib.import_module('.'+className, package=__name__+'.classes' + namespace)

            classPackage = ''
            for d in self.aicConfig['plugins']:
                # for t in ['tool', 'function']:
                path = 'plugin\\'+d+'\\'+namespace
                # '+t+'-class/
                pythonFile = os.path.dirname(os.path.realpath(__file__)) + '\\' + path + '\\' + className + '.py'
                if os.path.isfile(pythonFile):
                    classPackage = 'aicells-server.' + path.replace('\\', '.')
                    classFile = pythonFile

            if classPackage == '':
                raise Exception('AICFactory: class not found: ' + classNameWithNamespace)

            # module = importlib.import_module('.'+className, package='aicells-server.classes' + namespace)
            module = importlib.import_module('.'+className, package=classPackage)

            classType = getattr(module, className)

            classType.factory = self
            classType.classFile = classFile

            self.classTypeCache[classNameWithNamespace] = classType

        classInstance = classType(**args)
        # classInstance.factory = self
        # if classPackage != '':
        #     classInstance.classFile = classFile

        if hasattr(classInstance, "Init") and callable(classInstance.Init):
            classInstance.Init(**args)
        return classInstance
