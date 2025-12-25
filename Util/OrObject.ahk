#Requires AutoHotkey v2.0

/************************************************************************
 * @class OrObject
 * @brief Modificación de la clase Object que permite iterar
 * por sus propiedades en orden de creación.
 * 
 * Se utiliza como un objeto normal.
 * 
 * La única desventaja es que no se puede inicializar directamente con {}.
 * 
 * @author bitasuperactive
 * @date 19/12/2025
 * @version 1.0.0
 * @extends Object
 * @see https://www.autohotkey.com/docs/v2/lib/Object.htm
 ***********************************************************************/
class OrObject extends Object
{
    /**
     * @private
     * Nombres de las propiedades en orden de creación.
     */
    __props := [] ;

    /**
     * @public
     * Crea un nuevo objeto cuyas propiedades serán indexadas.
     * @param {Any} props Cada propiedad se asigna con 2 parámetros,
     * un String que será el nombre, y cualquier tipo para el valor.
     * @returns {OrObject}
     */
    __New(props*)
    {
        evenProps := Mod(props.Length, 2) = 0
        if (!evenProps)
            throw ValueError("Propiedades inválidas.")
        
        for key in props {
            evenIndex := Mod(A_Index, 2) = 0
            if (evenIndex)
                continue
            if (Type(key) != "String" || InStr(key, ' '))
                throw ValueError("Propiedades inválidas.")

            val := props[A_Index + 1]
            this.%key% := val
        }
        
        return this
    }

    /** @public */
    DefineProp(name, desc)
    {
        this.__props.Push(name)
        return super.DefineProp(name, desc)
    }

    /** @public */
    DeleteProp(name)
    {
        index := this.__HasProp(name)
        if (index > 0)
            this.__props.RemoveAt(index)
        
        return super.DeleteProp(name)
    }

    /** @public */
    OwnProps()
    {
        props := this.__props.Clone()
        return (&k := "", &v := "") => __Iterate(&k, &v)


        __Iterate(&k, &v)
        {
            if (A_Index > props.Length)
                return false

            k := props[A_Index]
            v := this.%k%
            return true
        }
    }

    /** @private */
    __Set(name, params, value)
    {
        if (name = "__props")
            super.DefineProp(name, { Value: value })
        else
            this.DefineProp(name, { Value: value })
        
        return value
    }

    /**
     * @private
     * Comprueba si la propiedad está definida en el array.
     * @param {String} porp Nombre de la propiedad objetivo.
     * @returns {Integer} Índice de la propiedad, o 0 si no la encuentra.
     */
    __HasProp(prop)
    {
        for p in this.__props {
            if (p = prop)
                return A_Index
        }
        return 0
    }
}