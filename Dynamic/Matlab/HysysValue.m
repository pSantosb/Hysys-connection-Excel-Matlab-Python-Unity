classdef HysysValue
    %HYSYSVALUE Summary of this class goes here
    %   Detailed explanation goes here
    
    properties
        insideObj
        isFlexVar
        index
        value
    end
    
    methods
        function obj = HysysValue(insideObj,isFlexVar, index)
            %HYSYSVALUE Construct an instance of this class
            %   Detailed explanation goes here
            if nargin<3
                index=-1;
            end
            if nargin<2
                isFlexVar=false;
            end
            obj.insideObj = insideObj;
            obj.isFlexVar=isFlexVar;
            obj.index=index;
        end
        
        function obj=set.value(obj,val)
            if obj.isFlexVar
                if obj.index==-1
                    obj.insideObj.Values=val;
                else
                    temp=obj.insideObj.Values;
                    temp(obj.index)=val;
                    obj.insideObj.Values=temp;
                end
            else
                obj.insideObj.Value =val;
            end
        end
        function v=get.value(obj)
            if obj.isFlexVar
                if obj.index==-1
                    v= obj.insideObj.Values;
                else
                    v= obj.insideObj.Values(obj.index);
                end
            else
                v= obj.insideObj.Value;
            end
        end
    end
end

