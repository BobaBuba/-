# -
Софт принимает зип архив с выписками на любые единицы кадастрового деления и выдает эксель файл с полной "матрешкой" - сначала вы увидите земельные участки, затем сооружения и здания внутри этих участков (по оформлению файла будет просто ориентироваться - везде есть отступы для простой визуализации), далее помещения и машино-места внутри. 

Площади являются барьерной метрикой в скрипте - метры меньшей кадастровой единицы не могут превышать метры большей (почти везде они будут равны, за исключением относительно редких случаев сооружений - теплострасс/инжереных сетей/другой инфраструктуры)
